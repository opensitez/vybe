! vybe-test: fortran/transfer_extended/transfer_scalar_minus_one
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer :: i = -1
if ((transfer(i, 0)) /= -1) then
    print *, "FAIL: want [-1] got [", transfer(i, 0), "]"
    stop 1
end if
end program t
