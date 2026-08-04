! vybe-test: fortran/matmul_dot_extended/dot_product_zero_first_operand
! origin: languages/fortran/tests/fortran/test_matmul_dot_extended.rs
program t
integer :: a(3) = [0, 0, 0]
integer :: b(3) = [2, 3, 4]
if ((dot_product(a, b)) /= 0) then
    print *, "FAIL: want [0] got [", dot_product(a, b), "]"
    stop 1
end if
end program t
