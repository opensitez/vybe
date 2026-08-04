! vybe-test: fortran/transfer_extended/transfer_int_real_roundtrip_one
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer :: i = 1, j
real :: r
r = transfer(i, 0.0)
j = transfer(r, 0)
if ((j) /= 1) then
    print *, "FAIL: want [1] got [", j, "]"
    stop 1
end if
end program t
