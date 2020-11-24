# oletools_dll
This *very* experimental project aims to produce a DLL (for Windows) 
to run some [oletools](https://github.com/decalage2/oletools) 
functions from any language other than Python, such as C or Golang.
This can be used for example to scan suspicious documents to detect
VBA macros and extract their source code, as it can be done in Python
with [olevba](https://github.com/decalage2/oletools/wiki/olevba).

For now the DLL is very basic: it only provides a single function that
takes a filename as argument, and returns a string containing the source
code of all VBA macros present in the file. There is no error handling yet.

In the background, oletools.dll loads the Python engine DLL and runs a
Python script that calls the [olevba API](https://github.com/decalage2/oletools/wiki/olevba#how-to-use-olevba-in-python-applications) 
from oletools.

A sample C client is also provided, to show how the DLL can be called from C.

# How it works

The oletools DLL is compiled from Python code thanks to 
[cffi](https://cffi.readthedocs.io/en/latest/), using its 
[embedding features](https://cffi.readthedocs.io/en/latest/embedding.html).

This is implemented using 3 files:
- oletools_dll_api.py implement the API of oletools.dll in python functions,
  which call oletools.
- oletools_dll.h defines the C API of oletools.dll, matching oletools_dll_api.py
- build_oletools_dll.py calls cffi to compile and build oletools.dll

# Quick demo

To test it, you may try the pre-built oletools.dll and the sample client
call_olevba.exe available in the repository:
1. Install **Python 3.9 64 bits** if you don't already have it 
   (other versions will not work with the pre-built DLL, 
   see below to build it yourself)
2. Install **oletools**: pip install -U oletools
3. Download **oletools.dll** and **call_olevba.exe** from the 
   [releases page](https://github.com/decalage2/oletools_dll/releases/tag/v0.0.1-alpha)
4. Copy both files to the same folder
5. In a CMD window, run `call_olevba.exe <filename>`, with `<filename>` pointing to a MS Office file
  with VBA macros.
6. the output should be similar to this:

```
c:\Users\xyz\Dev\oletools_dll\sample_client_C>call_olevba.exe Word_macro.doc
Sample C Client for the oletools DLL

Loading oletools.dll
Calling get_all_macros("Word_macro.doc"):
--- VBA CODE: -----------------------------------------------------------------
Attribute VB_Name = "ThisDocument"
Attribute VB_Base = "1Normal.ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = True
Attribute VB_Customizable = True
Attribute VB_Control = "CommandButton1, 0, 0, MSForms, CommandButton"
Private Sub CommandButton1_Click()
x = MsgBox("This is a VBA macro")
End Sub
-------------------------------------------------------------------------------
```  
  

# Requirements

To build the DLL (and optionally the sample C client), you will need:
- Python 3.x (tested with Python 3.9 64 bits)
- cffi (pip install -U cffi)
- a C compiler, such as Build Tools for Visual Studio 2019 -
  See https://wiki.python.org/moin/WindowsCompilers

To use the DLL, you will need:
- Python 3.x installed (same version as for the build)
- oletools installed (see [install instructions](https://github.com/decalage2/oletools/wiki/Install))
- oletools.dll in the same directory as the client, or reachable by PATH

# How to build the DLL

- download the files from this repository
- open a CMD window, go to the folder oletools_dll
- run python build_oletools_dll.py
- if everything goes well, oletools.dll should appear in the same directory

# How to build the sample C client

- if you use the Build Tools for Visual Studio, open a Visual C++ command 
  prompt for 64 bits
- go to the sample_client_C folder
- run cl call_olevba.c
- if everything goes well, call_olevba.exe should appear in the same directory

# How to run the sample C client

- copy oletools.dll in the same directory, or make sure it is reachable by PATH
- run `call_olevba.exe <filename>`, with `<filename>` pointing to a MS Office file
  with VBA macros.

# How to implement your own client

You should be able to call the oletools DLL from any language that can
load DLLs. The API of oletools.dll is defined in 
[oletools_dll.h](https://github.com/decalage2/oletools_dll/blob/main/oletools_dll/oletools_dll.h).

You can use the code of the 
[sample C client](https://github.com/decalage2/oletools_dll/blob/main/sample_client_C/call_olevba.c) 
as inspiration.