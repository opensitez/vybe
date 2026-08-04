! vybe-test: fortran/reshape_pad_extended/reshape_compile_empty_source_with_pad
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs

program t
    integer :: a(0)
    integer :: m(2)
    m = reshape(a, [2], pad=9)
    print *, m(1)
end program t
