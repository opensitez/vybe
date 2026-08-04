! vybe-test: fortran/transfer/transfer_size_expand
! origin: languages/fortran/tests/fortran/test_transfer.rs

program test
    integer :: a = 42
    integer :: b(4)
    b = transfer(a, b, 4)
    if ((b(1)) /= 42) then
    print *, "FAIL: want [42] got [", b(1), "]"
    stop 1
end if
    if ((b(2)) /= 0) then
    print *, "FAIL: want [0] got [", b(2), "]"
    stop 1
end if
    if ((b(3)) /= 0) then
    print *, "FAIL: want [0] got [", b(3), "]"
    stop 1
end if
end program test
