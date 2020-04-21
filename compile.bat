:: app name: RevisionZero
:: Set this file for compiling the executable of the macro.
:: So it can be added to the user custom theme in solidedge. 
ipyc.exe /main:./revzero/__main__.py ^
./revision_zero/Interop.SolidEdge.dll ^
./revision_zero/api.py ^
/embed ^
/out:revision_zero_macro_64x_0-0-0 ^
/platform:x64 ^
/standalone ^
/target:exe ^
/win32icon:logo.ico 
