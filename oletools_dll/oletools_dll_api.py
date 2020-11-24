"""
API for the oletools DLL

This script is loaded by build_oletools_dll.py to produce oletools.dll
It is not meant to be run directly.

This script defines and implements the API of oletools.dll in Python
See oletools_dll.h for the corresponding C API
"""

# TODO: add API using UTF-8 strings as input and output
# TODO: add error handling
# TODO: add function to check if there are macros

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

from oletools_c import ffi
from oletools.olevba import VBA_Parser

@ffi.def_extern()
def get_all_macros(filename_wchar):
    """
    Open a file and extract the source code of all VBA modules
    by calling oletools.VBA_Parser.get_vba_code_all_modules()
    :param filename_wchar: file to be opened (unicode wchar_t string)
    :return: source code of all VBA modules (unicode wchar_t string)
    """
    # print(wchar_string)
    filename_str = ffi.string(filename_wchar)
    print("filename: {}".format(filename_str))
    ovba = VBA_Parser(filename_str)
    # TODO: check if there are macros
    vba_code_str = ovba.get_vba_code_all_modules()
    # convert python str to C string
    return ffi.new("wchar_t[]", vba_code_str)
