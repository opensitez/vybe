! vybe-test: fortran/matmul_dot_extended/dot_product_slice_stride_two
! origin: languages/fortran/tests/fortran/test_matmul_dot_extended.rs
program t
integer :: a(7) = [1, 2, 3, 4, 5, 6, 7]
integer :: b(7) = [7, 6, 5, 4, 3, 2, 1]
if ((dot_product(a(1:7:2), b(1:7:2))) /= 44) then
    print *, "FAIL: want [44] got [", dot_product(a(1:7:2), b(1:7:2)), "]"
    stop 1
end if
end program t
