! vybe-test: fortran/transfer/transfer_array_roundtrip_and_slice
! origin: languages/fortran/tests/fortran/test_transfer.rs

program test
    integer :: source(2) = [11, 22]
    integer :: target(2)
    target = transfer(source, target)
    if ((target(1)) /= 11) then
    print *, "FAIL: want [11] got [", target(1), "]"
    stop 1
end if
    if ((target(2)) /= 22) then
    print *, "FAIL: want [22] got [", target(2), "]"
    stop 1
end if
end program test
