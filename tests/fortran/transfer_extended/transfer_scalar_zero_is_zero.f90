! vybe-test: fortran/transfer_extended/transfer_scalar_zero_is_zero
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer :: i = 0
if ((transfer(i, 0)) /= 0) then
    print *, "FAIL: want [0] got [", transfer(i, 0), "]"
    stop 1
end if
end program t
