! vybe-test: fortran/fortran2018_extended/sort_matrix_with_mask
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs

program t
    integer :: a(6) = [5, 2, 8, 1, 9, 3]
    logical :: mask(6) = [.true., .false., .true., .false., .true., .false.]
    call sort(a, mask=mask)
    print *, a(1)
end program t
