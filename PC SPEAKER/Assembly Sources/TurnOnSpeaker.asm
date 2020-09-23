;I used NASM
;the frequency is passed as parameter

%Define frequency [ebp+8]


[BITS 32]

push ebp
mov ebp,esp
push ebx    ; registers are preserved
push esi    ; 
push edi    ; 
  

mov ax,0b6h          
out 43h,ax         	;Preparing timer by sending 10111100 to port 43.
mov eax,1193180
mov edx,0
mov ebx,[ebp+8]
div ebx             ;Divide the frequency by timer ticks per second
mov ebx,eax         ;and write byte to byte to timer
xor ax,ax
movzx ax,bl
out 42h,ax              
xor ax,ax
movzx ax,bh
out 42h,ax
in ax,61h        ;Save speaker control byte
xor cx,cx
mov cx,ax
or ax,3
out 61h,ax
xor eax,eax
movzx eax,cx     ;Send the speaker control byte as return value




pop edi    ; Restore register values
pop esi
pop ebx
mov esp,ebp
pop ebp
ret 16