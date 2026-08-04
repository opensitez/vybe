! vybe-test: fortran/transfer/transfer_size_expands_with_zero_padding
! origin: languages/fortran/tests/fortran/test_transfer.rs

program test
    integer :: source
    integer :: target(4)
    source = 99
    target = transfer(source, target, 4)
    if ((target(1)) /= 99) then
    print *, "FAIL: want [99] got [", target(1), "]"
    stop 1
end if
    if ((target(2)) /= 0) then
    print *, "FAIL: want [0] got [", target(2), "]"
    stop 1
end if
    if ((target(3)) /= 0) then
    print *, "FAIL: want [0] got [", target(3), "]"
    stop 1
end if
    if ((target(4)) /= 0) then
    print *, "FAIL: want [0] got [", target(4), "]"
    stop 1
end if
end program test
