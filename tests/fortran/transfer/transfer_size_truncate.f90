! vybe-test: fortran/transfer/transfer_size_truncate
! origin: languages/fortran/tests/fortran/test_transfer.rs

program test
    integer :: a(4) = [1, 2, 3, 4]
    integer :: b(2)
    b = transfer(a, b, 2)
    if ((b(1)) /= 1) then
    print *, "FAIL: want [1] got [", b(1), "]"
    stop 1
end if
end program test
