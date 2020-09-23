.486                                    ;# Create 32 bit code
.model flat, stdcall                    ;# 32 bit memory model
option casemap :none                    ;# Case sensitive
include Macros.inc                      ;# Macros 'n stuff

.code
start:
MyProc proc     
    LOCAL   nTemp   :DWORD          

    push    ebx    

;Debug
;   int     3

;Check that the CPUID instruction is supported
    pushfd
    pop	eax
    btc	eax, 15h
    push	eax
    popfd
    pushfd
    pop	edx
    xor	eax, edx
    jz	_cpuid_supported
   
    xor     eax, eax
    xor     ebx, ebx
    xor     ecx, ecx
    xor     edx, edx
    jmp     _set_results

_cpuid_supported:
    mov     eax,[ebp+0Ch]
    db      0Fh, 0A2h                       ;# CPUID

_set_results:
    mov     nTemp, edx
    
    mov     edx,[ebp+10h]
    mov     [edx], eax

    mov     edx,[ebp+14h]
    mov     [edx], ebx

    mov     edx,[ebp+18h]
    mov     [edx], ecx

    mov     edx,[ebp+1Ch]
    mov     eax, nTemp
    mov     [edx], eax

    pop     ebx
    xor     eax, eax
    ret     18h
MyProc endp
end start