! vybe-test: fortran/arrays/intrinsic_minval
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    integer :: a(5) = [3, 1, 4, 1, 5]
    print *, minval(a)
end program test
