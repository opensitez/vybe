! vybe-test: fortran/matmul_dot_extended/dot_product_both_zero_length_four
! origin: languages/fortran/tests/fortran/test_matmul_dot_extended.rs
program t
integer :: a(4) = [0, 0, 0, 0]
integer :: b(4) = [0, 0, 0, 0]
if ((dot_product(a, b)) /= 0) then
    print *, "FAIL: want [0] got [", dot_product(a, b), "]"
    stop 1
end if
end program t
