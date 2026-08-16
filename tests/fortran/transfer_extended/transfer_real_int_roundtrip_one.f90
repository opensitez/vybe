! vybe-test: fortran/transfer_extended/transfer_real_int_roundtrip_one
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
real :: x = 1.0
integer :: n
n = transfer(x, 0)
if (.not. (transfer(n, 0.0) == x)) then
    print *, "FAIL: want [1] got [", transfer(n, 0.0) == x, "]"
    stop 1
end if
end program t
