! vybe-test: fortran/transfer_extended/transfer_int_to_real_then_back_equality
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer :: i = 12345, j
real :: r
r = transfer(i, 0.0)
j = transfer(r, 0)
if (.not. (i == j)) then
    print *, "FAIL: want [1] got [", i == j, "]"
    stop 1
end if
end program t
