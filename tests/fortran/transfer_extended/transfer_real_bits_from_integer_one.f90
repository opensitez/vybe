! vybe-test: fortran/transfer_extended/transfer_real_bits_from_integer_one
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer :: i = 1
real :: r
r = transfer(i, 0.0)
if ((transfer(r, 0)) /= 1) then
    print *, "FAIL: want [1] got [", transfer(r, 0), "]"
    stop 1
end if
end program t
