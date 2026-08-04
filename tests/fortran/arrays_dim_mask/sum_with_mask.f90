! vybe-test: fortran/arrays_dim_mask/sum_with_mask
! origin: languages/fortran/tests/fortran/test_arrays_dim_mask.rs

program test
    integer :: a(6) = [1, 2, 3, 4, 5, 6]
    logical :: mask(6) = [.true., .false., .true., .false., .true., .false.]
    print *, sum(a, mask=mask)
end program test
