! vybe-test: fortran/transfer/transfer_complex_to_real_pair
! origin: languages/fortran/tests/fortran/test_transfer.rs

program test
    complex :: c = (1.0, 2.0)
    real :: pair(2)
    pair = transfer(c, pair)
    if ((pair(1)) /= 1) then
    print *, "FAIL: want [1] got [", pair(1), "]"
    stop 1
end if
    if ((pair(2)) /= 2) then
    print *, "FAIL: want [2] got [", pair(2), "]"
    stop 1
end if
end program test
