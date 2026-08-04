! vybe-test: fortran/legacy_data_extended/data_implied_do_with_step
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

program t
    integer :: a(5)
    data (a(i), i = 1, 5, 2) /10, 20, 30/
    print *, a(1), a(3), a(5)
end program t
