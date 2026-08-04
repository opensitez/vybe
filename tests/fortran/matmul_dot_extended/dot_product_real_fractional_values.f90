! vybe-test: fortran/matmul_dot_extended/dot_product_real_fractional_values
! origin: languages/fortran/tests/fortran/test_matmul_dot_extended.rs
program t
real :: a(3) = [0.5, 1.5, 2.0]
real :: b(3) = [2.0, 2.0, 1.0]
if ((dot_product(a, b)) /= 6) then
    print *, "FAIL: want [6] got [", dot_product(a, b), "]"
    stop 1
end if
end program t
