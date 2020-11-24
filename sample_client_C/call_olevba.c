/* Sample C client for the oletools DLL

How to compile and build:
- requires a C compiler, such as Build Tools for Visual Studio 2019
  See https://wiki.python.org/moin/WindowsCompilers
- Open Visual C++ 64 bits command prompt
- cl call_olevba.c

Usage:
call_olevba <filename>

Requires:
- Python 3.9 64 bits installed
- oletools.dll in the current directory or reachable by PATH
- oletools installed (pip install -U oletools)


Copyright (c) 2020, Philippe Lagadec
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this
   list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice,
   this list of conditions and the following disclaimer in the documentation
   and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

*/
#include <stdio.h>
#include <stdlib.h>
#include <windows.h>

#ifndef CFFI_DLLEXPORT
#  if defined(_MSC_VER)
#    define CFFI_DLLEXPORT  extern __declspec(dllimport)
#  else
#    define CFFI_DLLEXPORT  extern
#  endif
#endif

//CFFI_DLLEXPORT int do_stuff(int, int);

typedef wchar_t * (*get_all_macros_type)(wchar_t *);
get_all_macros_type get_all_macros;

//int main()
wmain( int argc, wchar_t *argv[ ], wchar_t *envp[ ] )
{
    HANDLE ldll;
    wchar_t *filename;

    printf("Sample C Client for the oletools DLL\n\n");
    if (argc != 2)
    {
        printf("Usage: call_olevba <file>\n");
        return -1;
    }

    printf("Loading oletools.dll\n");
    ldll = LoadLibrary("oletools.dll");
    if(ldll>(void*)HINSTANCE_ERROR)
    {
        get_all_macros = (get_all_macros_type) GetProcAddress(ldll, "get_all_macros");
        filename = argv[1];
        printf("Calling get_all_macros(\"%S\"):\n", filename);
        printf("--- VBA CODE: -----------------------------------------------------------------\n");
        printf("%S\n", get_all_macros(filename));
        printf("-------------------------------------------------------------------------------\n");
    }
    else
    {
        printf("ERROR while loading oletools.dll.\n");
        printf("Please make sure that:\n");
        printf("1) oletools.dll is in the current directory, or reachable by PATH.\n");
        printf("2) Python 3.9 64 bits is installed.\n");
        printf("3) oletools is installed (pip install -U oletools).\n");

        //int error = GetLastError()
    }

}