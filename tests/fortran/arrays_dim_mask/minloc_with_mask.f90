! vybe-test: fortran/arrays_dim_mask/minloc_with_mask
! origin: languages/fortran/tests/fortran/test_arrays_dim_mask.rs

program test
    integer :: a(5) = [5, 1, 4, 1, 5]
    logical :: mask(5) = [.true., .false., .true., .true., .true.]
    integer :: loc(1)
    loc = minloc(a, mask=mask)
    print *, loc(1)
end program test
