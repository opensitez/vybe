! vybe-test: fortran/transfer_extended/transfer_logical_scalar_to_integer
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer :: true_bits
integer :: false_bits
true_bits = transfer(.true., 0)
false_bits = transfer(.false., 0)
if ((true_bits) /= 1) then
    print *, "FAIL: want [1] got [", true_bits, "]"
    stop 1
end if
if ((false_bits) /= 0) then
    print *, "FAIL: want [0] got [", false_bits, "]"
    stop 1
end if
end program t
