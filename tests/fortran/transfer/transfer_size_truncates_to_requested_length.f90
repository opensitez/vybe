! vybe-test: fortran/transfer/transfer_size_truncates_to_requested_length
! origin: languages/fortran/tests/fortran/test_transfer.rs

program test
    integer :: source(4) = [10, 20, 30, 40]
    integer :: target(2)
    target = transfer(source, target, 2)
    if ((target(1)) /= 10) then
    print *, "FAIL: want [10] got [", target(1), "]"
    stop 1
end if
    if ((target(2)) /= 20) then
    print *, "FAIL: want [20] got [", target(2), "]"
    stop 1
end if
end program test
