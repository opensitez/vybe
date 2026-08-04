! vybe-test: fortran/transfer_extended/transfer_char_a_to_integer_le
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
character(len=1) :: c = 'A'
if ((transfer(c, 0)) /= 65) then
    print *, "FAIL: want [65] got [", transfer(c, 0), "]"
    stop 1
end if
end program t
