! vybe-test: fortran/arrays_dim_mask/maxloc_with_mask
! origin: languages/fortran/tests/fortran/test_arrays_dim_mask.rs

program test
    integer :: a(6) = [1, 9, 2, 8, 3, 7]
    logical :: mask(6) = [.false., .false., .true., .true., .true., .true.]
    integer :: loc(1)
    loc = maxloc(a, mask=mask)
    print *, loc(1)
end program test
