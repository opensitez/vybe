! vybe-test: fortran/matmul_dot_extended/dot_product_negative_mixed_signs
! origin: languages/fortran/tests/fortran/test_matmul_dot_extended.rs
program t
integer :: a(3) = [-1, 2, -3]
integer :: b(3) = [4, -5, 6]
if ((dot_product(a, b)) /= -32) then
    print *, "FAIL: want [-32] got [", dot_product(a, b), "]"
    stop 1
end if
end program t
