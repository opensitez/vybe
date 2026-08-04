! vybe-test: fortran/transfer/transfer_array_to_scalar
! origin: languages/fortran/tests/fortran/test_transfer.rs

program test
    integer(kind=4) :: parts(2) = [0, 0]
    integer(kind=8) :: big
    big = transfer(parts, 0_8)
    if ((big == 0) /= 1) then
    print *, "FAIL: want [1] got [", big == 0, "]"
    stop 1
end if
end program test
