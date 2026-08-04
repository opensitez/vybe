! vybe-test: fortran/transfer/transfer_with_size
! origin: languages/fortran/tests/fortran/test_transfer.rs

program test
    integer :: a(4) = [0, 0, 0, 0]
    integer(kind=1) :: bytes(16)
    bytes = transfer(a, bytes)
    if ((bytes(1)) /= 0) then
    print *, "FAIL: want [0] got [", bytes(1), "]"
    stop 1
end if
    if ((bytes(2)) /= 0) then
    print *, "FAIL: want [0] got [", bytes(2), "]"
    stop 1
end if
    if ((bytes(3)) /= 0) then
    print *, "FAIL: want [0] got [", bytes(3), "]"
    stop 1
end if
end program test
