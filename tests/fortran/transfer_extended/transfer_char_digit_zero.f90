! vybe-test: fortran/transfer_extended/transfer_char_digit_zero
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
character(len=1) :: c = '0'
if ((transfer(c, 0)) /= 48) then
    print *, "FAIL: want [48] got [", transfer(c, 0), "]"
    stop 1
end if
end program t
