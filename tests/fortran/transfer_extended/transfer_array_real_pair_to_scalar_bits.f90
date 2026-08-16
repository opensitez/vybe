! vybe-test: fortran/transfer_extended/transfer_array_real_pair_to_scalar_bits
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
real :: pair(2) = [1.0, 2.0]
integer :: n
n = transfer(pair, 0)
if (.not. (transfer(n, pair(1)) == 1.0)) then
    print *, "FAIL: want [1] got [", transfer(n, pair(1)) == 1.0, "]"
    stop 1
end if
end program t
