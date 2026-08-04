! vybe-test: fortran/implied_do_forms/data_implied_do_second_matrix_row
! origin: languages/fortran/tests/fortran/test_implied_do_forms.rs

program t
    integer :: grid(2, 3)
    data (grid(2, j), j = 1, 3) /10, 20, 30/
    print *, grid(2, 2)
end program t
