! vybe-test: fortran/transfer/transfer_int_to_real
! origin: languages/fortran/tests/fortran/test_transfer.rs

program test
    integer :: i = 0
    real :: r
    r = transfer(i, 0.0)
    if ((r) /= 0) then
    print *, "FAIL: want [0] got [", r, "]"
    stop 1
end if
end program test
