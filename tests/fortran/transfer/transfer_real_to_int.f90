! vybe-test: fortran/transfer/transfer_real_to_int
! origin: languages/fortran/tests/fortran/test_transfer.rs

program test
    real :: x = 0.0
    integer :: i
    i = transfer(x, 0)
    if ((i) /= 0) then
    print *, "FAIL: want [0] got [", i, "]"
    stop 1
end if
end program test
