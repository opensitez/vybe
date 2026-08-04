! vybe-test: fortran/implied_do_forms/data_implied_do_nested_indices_two_dimensions
! origin: languages/fortran/tests/fortran/test_implied_do_forms.rs

program t
    integer :: mat(2, 3)
    data ((mat(i, j), j = 1, 3), i = 1, 2) /1, 2, 3, 4, 5, 6/
    print *, mat(2, 3)
end program t
