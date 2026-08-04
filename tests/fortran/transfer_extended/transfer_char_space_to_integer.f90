! vybe-test: fortran/transfer_extended/transfer_char_space_to_integer
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
character(len=1) :: c = ' '
if ((transfer(c, 0)) /= 32) then
    print *, "FAIL: want [32] got [", transfer(c, 0), "]"
    stop 1
end if
end program t
