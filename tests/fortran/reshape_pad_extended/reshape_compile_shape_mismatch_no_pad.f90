! vybe-test: fortran/reshape_pad_extended/reshape_compile_shape_mismatch_no_pad
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs

program t
    integer :: a(3) = [1, 2, 3]
    integer :: m(2,2)
    m = reshape(a, [2, 2])
    print *, m(1,1)
end program t
