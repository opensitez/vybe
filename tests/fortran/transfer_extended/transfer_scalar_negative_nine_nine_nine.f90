! vybe-test: fortran/transfer_extended/transfer_scalar_negative_nine_nine_nine
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer :: i = -999
if ((transfer(i, 0)) /= -999) then
    print *, "FAIL: want [-999] got [", transfer(i, 0), "]"
    stop 1
end if
end program t
