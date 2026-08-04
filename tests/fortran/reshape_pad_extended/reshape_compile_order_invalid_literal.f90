! vybe-test: fortran/reshape_pad_extended/reshape_compile_order_invalid_literal
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs

program t
    integer :: a(4) = [1, 2, 3, 4]
    integer :: m(2,2)
    m = reshape(a, [2, 2], order='X')
    print *, m(1,1)
end program t
