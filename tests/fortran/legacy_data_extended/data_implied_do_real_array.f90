! vybe-test: fortran/legacy_data_extended/data_implied_do_real_array
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

program t
    real :: r(4)
    data (r(i), i = 1, 4) /1.5, 2.5, 3.5, 4.5/
    print *, r(2)
end program t
