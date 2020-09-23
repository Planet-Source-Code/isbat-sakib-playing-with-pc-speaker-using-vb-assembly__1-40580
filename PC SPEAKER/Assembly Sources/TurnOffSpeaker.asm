;I used NASM

%Define frequency [ebp+8]


[BITS 32]

push ebp
mov ebp,esp
push ebx    ; registers are preserved
push esi    
push edi    


mov eax,[ebp+8]    ;getting the value passed as parameter in eax
out 61h,ax         ;then sending this to port 61h



pop edi 		   ; Restore register values
pop esi
pop ebx
mov esp,ebp
pop ebp
ret 16