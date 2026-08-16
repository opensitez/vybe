! vybe-test: fortran/reshape_pad_extended/reshape_compile_3d_negative_extent
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
! vybe-test-mode: compile-fail

program t
    integer :: a(8) = [(i, i = 1, 8)]
    integer :: m(2,2,-1)
    m = reshape(a, [2, 2, -1])
    print *, 0
end program t
