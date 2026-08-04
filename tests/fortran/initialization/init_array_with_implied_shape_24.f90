! vybe-test: fortran/initialization/init_array_with_implied_shape_24
! origin: languages/fortran/tests/fortran/test_initialization.rs
program p
integer, dimension(2,3) :: m = reshape([1,2,3,4,5,6], [2,3])
print *, m(2,2)
end program p
