! vybe-test: fortran/arrays_dim_mask/maxval_with_mask
! origin: languages/fortran/tests/fortran/test_arrays_dim_mask.rs

program test
    integer :: a(6) = [1, 9, 2, 8, 3, 7]
    logical :: mask(6) = [.false., .false., .true., .true., .true., .true.]
    print *, maxval(a, mask=mask)
end program test
