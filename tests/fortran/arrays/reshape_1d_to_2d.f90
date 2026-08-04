! vybe-test: fortran/arrays/reshape_1d_to_2d
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    integer :: a(6) = [1, 2, 3, 4, 5, 6]
    integer :: m(2,3)
    m = reshape(a, [2, 3])
    print *, m(1,1)
end program test
