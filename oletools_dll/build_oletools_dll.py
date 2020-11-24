"""
Script to build the oletools DLL using CFFI

Usage: python build_oletools_dll.py

Requires:
- Python 3.x
- cffi (pip install -U cffi)
- a C compiler, such as Build Tools for Visual Studio 2019
  See https://wiki.python.org/moin/WindowsCompilers
- oletools_dll.h to define the C API of the oletools.dll
- oletools_dll_api.py to implement the API in Python
"""

# Copyright (c) 2020, Philippe Lagadec
# All rights reserved.
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
# 1. Redistributions of source code must retain the above copyright notice, this
#    list of conditions and the following disclaimer.
#
# 2. Redistributions in binary form must reproduce the above copyright notice,
#    this list of conditions and the following disclaimer in the documentation
#    and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
# AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
# IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
# FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
# DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
# SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
# CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
# OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
# OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.


import cffi
ffibuilder = cffi.FFI()

# DLL functions
with open('oletools_dll.h') as f:
    oletools_dll_h = f.read()
ffibuilder.embedding_api(oletools_dll_h)

# Name of intermediate C file
ffibuilder.set_source("oletools_c", "")

# Python code for DLL
with open('oletools_dll_api.py') as f:
    api_script = f.read()
ffibuilder.embedding_init_code(api_script)

# Compile DLL
ffibuilder.compile(target="oletools.*", verbose=True)