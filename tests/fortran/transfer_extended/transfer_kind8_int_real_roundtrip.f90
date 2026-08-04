! vybe-test: fortran/transfer_extended/transfer_kind8_int_real_roundtrip
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer(kind=8) :: i = 9876543210_8, j
real(kind=8) :: r
r = transfer(i, 0.0d0)
j = transfer(r, 0_8)
if ((j) /= 9876543210) then
    print *, "FAIL: want [9876543210] got [", j, "]"
    stop 1
end if
end program t
