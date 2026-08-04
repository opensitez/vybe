! vybe-test: fortran/reshape_pad_extended/reshape_compile_large_shape_vector
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs

program t
    integer :: a(6) = [1, 2, 3, 4, 5, 6]
    integer :: sh(4) = [1, 1, 2, 3]
    integer :: m(1,1,2,3)
    m = reshape(a, sh)
    print *, sum(m)
end program t
