; David Fritts corrected by Robert Rayment

.486                                    ;# Create 32 bit code
.model flat, stdcall                    ;# 32 bit memory model
option casemap :none                    ;# Case sensitive
include Macros.inc                      ;# Macros 'n stuff

.code
start:
shiftright proc

mov eax, dword ptr [esp+08h] ; value
mov ecx, dword ptr [esp+0Ch] ; shift
mov edx, dword ptr [esp+10h] ; hresult

shr eax, cl
mov dword ptr [edx], eax
xor eax, eax

ret 16
shiftright endp
end start
