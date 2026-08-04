! vybe-test: fortran/transfer/transfer_int_roundtrip_same_type
! origin: languages/fortran/tests/fortran/test_transfer.rs

program test
    integer :: original
    integer :: copy
    original = -17
    copy = transfer(original, 0)
    if ((copy) /= -17) then
    print *, "FAIL: want [-17] got [", copy, "]"
    stop 1
end if
    if ((original == copy) /= 1) then
    print *, "FAIL: want [1] got [", original == copy, "]"
    stop 1
end if
end program test
