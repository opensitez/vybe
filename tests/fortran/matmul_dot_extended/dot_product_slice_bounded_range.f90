! vybe-test: fortran/matmul_dot_extended/dot_product_slice_bounded_range
! origin: languages/fortran/tests/fortran/test_matmul_dot_extended.rs
program t
integer :: a(6) = [10, 20, 30, 40, 50, 60]
integer :: b(6) = [1, 1, 1, 1, 1, 1]
if ((dot_product(a(2:4), b(2:4))) /= 90) then
    print *, "FAIL: want [90] got [", dot_product(a(2:4), b(2:4)), "]"
    stop 1
end if
end program t
