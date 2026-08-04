! vybe-test: fortran/legacy_data_extended/data_implied_do_partial_matrix_row
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

program t
    integer :: table(3, 2)
    data (table(1, j), j = 1, 2) /7, 8/
    print *, table(1, 1)
end program t
