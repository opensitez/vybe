! vybe-test: fortran/transfer/transfer_logical_values_to_integer
! origin: languages/fortran/tests/fortran/test_transfer.rs

program test
    logical :: l1, l2
    integer :: bits(2)
    l1 = .true.
    l2 = .false.
    bits = transfer([l1, l2], bits)
    if ((bits(1)) /= 1) then
    print *, "FAIL: want [1] got [", bits(1), "]"
    stop 1
end if
    if ((bits(2)) /= 0) then
    print *, "FAIL: want [0] got [", bits(2), "]"
    stop 1
end if
end program test
