! vybe-test: fortran/arrays_dim_mask/findloc_with_mask
! origin: languages/fortran/tests/fortran/test_arrays_dim_mask.rs

program test
    integer :: a(6) = [1, 2, 1, 2, 1, 2]
    logical :: mask(6) = [.false., .true., .true., .true., .true., .true.]
    integer :: loc(1)
    loc = findloc(a, 1, mask=mask)
    print *, loc(1)
end program test
