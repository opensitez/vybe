! vybe-test: fortran/transfer/transfer_double_to_two_ints
! origin: languages/fortran/tests/fortran/test_transfer.rs

program test
    real(kind=8) :: d = 0.0d0
    integer :: parts(2)
    parts = transfer(d, parts)
    if ((parts(1)) /= 0) then
    print *, "FAIL: want [0] got [", parts(1), "]"
    stop 1
end if
    if ((parts(2)) /= 0) then
    print *, "FAIL: want [0] got [", parts(2), "]"
    stop 1
end if
end program test
