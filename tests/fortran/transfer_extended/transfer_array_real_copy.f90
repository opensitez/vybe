! vybe-test: fortran/transfer_extended/transfer_array_real_copy
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
real :: a(3) = [1.5, 2.5, 3.5]
real :: b(3)
b = transfer(a, b)
if (abs((b(2)) - 2.5) > 1.0e-6) then
    print *, "FAIL: want [2.5] got [", b(2), "]"
    stop 1
end if
end program t
