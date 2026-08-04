! vybe-test: fortran/legacy_data_extended/data_implied_do_multiple_sets
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

program t
    integer :: a(3), b(2)
    data (a(i), i = 1, 3) /1, 2, 3/, (b(j), j = 1, 2) /8, 9/
    print *, a(1) + b(2)
end program t
