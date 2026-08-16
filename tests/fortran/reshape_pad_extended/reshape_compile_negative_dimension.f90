! vybe-test: fortran/reshape_pad_extended/reshape_compile_negative_dimension
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
! vybe-test-mode: compile-fail

program t
    integer :: a(4) = [1, 2, 3, 4]
    integer :: m(-1,2)
    m = reshape(a, [-1, 2])
    print *, m(1,1)
end program t
