! vybe-test: fortran/transfer_extended/transfer_size_two_from_scalar
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer :: a = 99
integer :: b(2)
b = transfer(a, b, 2)
if ((b(1)) /= 99) then
    print *, "FAIL: want [99] got [", b(1), "]"
    stop 1
end if
end program t
