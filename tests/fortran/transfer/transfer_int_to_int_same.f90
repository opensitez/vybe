! vybe-test: fortran/transfer/transfer_int_to_int_same
! origin: languages/fortran/tests/fortran/test_transfer.rs

program test
    integer :: a = 42, b
    b = transfer(a, 0)
    if ((b) /= 42) then
    print *, "FAIL: want [42] got [", b, "]"
    stop 1
end if
end program test
