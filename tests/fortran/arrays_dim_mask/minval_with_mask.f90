! vybe-test: fortran/arrays_dim_mask/minval_with_mask
! origin: languages/fortran/tests/fortran/test_arrays_dim_mask.rs

program test
    integer :: a(6) = [10, 1, 20, 2, 30, 3]
    logical :: mask(6) = [.true., .false., .true., .false., .true., .false.]
    print *, minval(a, mask=mask)
end program test
