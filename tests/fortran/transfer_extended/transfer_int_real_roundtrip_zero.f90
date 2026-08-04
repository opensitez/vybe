! vybe-test: fortran/transfer_extended/transfer_int_real_roundtrip_zero
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer :: i = 0, j
real :: r
r = transfer(i, 0.0)
j = transfer(r, 0)
if ((j) /= 0) then
    print *, "FAIL: want [0] got [", j, "]"
    stop 1
end if
end program t
