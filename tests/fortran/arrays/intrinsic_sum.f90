! vybe-test: fortran/arrays/intrinsic_sum
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    print *, sum(a)
end program test
