! vybe-test: fortran/transfer_extended/transfer_size_expand_scalar_first_element
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer :: a = 42
integer :: b(4)
b = transfer(a, b, 4)
if ((b(1)) /= 42) then
    print *, "FAIL: want [42] got [", b(1), "]"
    stop 1
end if
end program t
