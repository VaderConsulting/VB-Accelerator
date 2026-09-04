                     FastCRC Library Version 1.50 ReadMe File
                                    July 2003
 
                    (C) Copyright 1999-2003 SlavaSoft Inc.

    Contents

    1. Warranty Disclaimer
    2. Introduction
    3. FileList	
    4. Installing and Uninstalling
    5. Trial Version Limitation


    1. Warranty Disclaimer
    ----------------------
    THE SOFTWARE IS PROVIDED AS IS WITHOUT WARRANTY OF ANY KIND. TO THE MAXIMUM EXTENT PERMITTED BY APPLICABLE LAW, SLAVASOFT INC. FURTHER DISCLAIMS ALL WARRANTIES, INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, AND NONINFRINGEMENT. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE PRODUCT AND DOCUMENTATION REMAINS WITH RECIPIENT. TO THE MAXIMUM EXTENT PERMITTED BY APPLICABLE LAW, IN NO EVENT SHALL SLAVASOFT INC. BE LIABLE FOR ANY CONSEQUENTIAL, INCIDENTAL, DIRECT, INDIRECT, SPECIAL, PUNITIVE, OR OTHER DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS, BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE LICENSE AGREEMENT OR THE USE OF OR INABILITY TO USE THE PRODUCT, EVEN IF SLAVASOFT INC. HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES.


    2. Introduction
    ---------------
    FastCRC Library is the fastest implementation of the most popular CRC algorithms. It allows Windows developers to perform in their applications checksum calculations for memory blocks, strings, blobs, streaming data and single files.
    FastCRC Library was designed to be fast, flexible and extremely easy to use. 
    FastCRC Library supports:

     3 checksum algorithms: CRC16, CRC16C, CRC32.

    FastCRC Library has 3 interfaces:

     1) FastCRC API ( The FastCRC API can be used as is, or it can be used to create wrappers for other languages that can use DLLs ).
     2) C++ Interface ( Wrapper classes to FastCRC API ).
     3) FastCRC Type Library ( Includes all declarations for accessing the FastCRC Library from different programming languages which have the possibility to reference a Type Library, such as VB 5.0, VB 6.0, VBA, etc. )

    FastCRC Library supports:

     1) Static Linking - When you perform static linking to FastCRC Library, the library code is included in your executable file the same way C functions are linked when you use a standard C library.
     2) Dynamic Linking - When you perform dynamic linking to  FastCRC Library, the library code  is loaded at run time from FastCRC.dll, and this DLL can be shared by multiple applications.  


    3. File List
    ------------
    FastCRC.h      - Include file for C/C++ applications
    FastCRC.lib    - LIB file for dynamic linking to FastCRC Library
    FastCRCS.lib   - LIB file for static linking to FastCRC Library
    FastCRC.dll    - DLL file    
    FastCRC.tlb    - type library for FastCRC.dll     
    FastCRC.chm    - help file
    ReadMe.txt       - this file
    Register.exe     - registration program for SlavaSoft products 

    4. Installing and Uninstalling
    ------------------------------
    To install FastCRC Library, perform the following:

    - If you want to work with FastCRC Library through its Type Library, follow these steps:

	1. Copy FastCRC.dll in one of the  executable paths for your operating system, or for your application.

	2. Copy FastCRC.tlb to any directory and add a reference to the FastCRC.tlb in your project(s).

    - If you want to work with FastCRC Library in C/C++, follow these steps:

	1. Copy FastCRC.dll in one of the  executable paths for your operating system, or for your application.

	2. Copy FastCRC.h in the include path for your project(s), or in the include path specified in your compiler settings.

	3. Copy FastCRC.lib and FastCRCS.lib in the library path for your project(s), or in the library path specified in your compiler settings. 

	4. Specify predefined symbols for your project:

	*** Define _SL_STATIC if you want to link to FastCRCS.lib (for static linking to FastCRC Library).

	*** Define _SL_NOFORCE_LIBS if you want to manually specify the necessary library for your project.

	Notes. 

	- Do not define any predefined symbols mentioned above if you want to link to FastCRC.lib (for dynamic linking to FastCRC Library).

	- Do define _SL_NOFORCE_LIBS if you use a compiler other than Microsoft Visual C++ (2.0 or later).

    In addition to the previous steps, perform the following:

	1. Copy the help files FastCRC.chm and ReadMe.txt in any directory where it will be easy for you to access them when FastCRC Library help will be needed.

	2. Copy the registration program Register.exe in any directory where it will be easy for you to access it when you will need to register FastCRC Library. 

    To uninstall FastCRC Library just remove all files listed above from your harddrive.


    5. Trial Version Limitation
    ---------------------------
    The trial version of FastCRC Library will cause a 10 seconds delay when running any application that uses it. To remove this delay you have to order the license for FastCRC Library and register it.

    