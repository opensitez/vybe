! vybe-test: fortran/transfer_extended/transfer_int_to_char_ichar_sixty_five
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer :: n = 65
character(len=1) :: c
 c = transfer(n, ' ')
if ((ichar(c)) /= 65) then
    print *, "FAIL: want [65] got [", ichar(c), "]"
    stop 1
end if
end program t
