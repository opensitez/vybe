! vybe-test: fortran/transfer_extended/transfer_array_int_in_expression
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer :: a(2) = [5, 6]
integer :: b(2)
b = transfer(a, b)
if ((b(1) + b(2)) /= 11) then
    print *, "FAIL: want [11] got [", b(1) + b(2), "]"
    stop 1
end if
end program t
