! vybe-test: fortran/reshape_pad_extended/reshape_compile_zero_dimension
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs

program t
    integer :: a(4) = [1, 2, 3, 4]
    integer :: m(0,2)
    m = reshape(a, [0, 2])
    print *, 0
end program t
