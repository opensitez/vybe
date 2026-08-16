! vybe-test: fortran/reshape_pad_extended/reshape_compile_pad_without_source
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs

program t
    integer :: m(3,3)
    m = reshape([(i, i = 1, 4)], [3, 3], pad=[0])
    print *, m(3,3)
end program t
