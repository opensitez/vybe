! vybe-test: fortran/legacy_data_extended/data_implied_do_descending_step
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

program t
    integer :: a(5)
    data (a(i), i = 5, 1, -1) /50, 40, 30, 20, 10/
    print *, a(5), a(1)
end program t
