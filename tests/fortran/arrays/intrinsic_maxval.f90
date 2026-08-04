! vybe-test: fortran/arrays/intrinsic_maxval
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    integer :: a(5) = [3, 1, 4, 1, 5]
    print *, maxval(a)
end program test
