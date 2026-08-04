! vybe-test: fortran/arrays_dim_mask/product_with_mask
! origin: languages/fortran/tests/fortran/test_arrays_dim_mask.rs

program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    logical :: mask(5) = [.true., .true., .false., .true., .false.]
    print *, product(a, mask=mask)
end program test
