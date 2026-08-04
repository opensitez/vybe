! vybe-test: fortran/transfer/transfer_scalar_to_array
! origin: languages/fortran/tests/fortran/test_transfer.rs

program test
    integer(kind=8) :: big = 0_8
    integer(kind=4) :: parts(2)
    parts = transfer(big, parts)
    if ((parts(1)) /= 0) then
    print *, "FAIL: want [0] got [", parts(1), "]"
    stop 1
end if
    if ((parts(2)) /= 0) then
    print *, "FAIL: want [0] got [", parts(2), "]"
    stop 1
end if
end program test
