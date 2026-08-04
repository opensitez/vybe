! vybe-test: fortran/arrays_dim_mask/count_basic_mask
! origin: languages/fortran/tests/fortran/test_arrays_dim_mask.rs

program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    print *, count(a > 3)
end program test
