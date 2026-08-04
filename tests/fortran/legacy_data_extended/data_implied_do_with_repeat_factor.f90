! vybe-test: fortran/legacy_data_extended/data_implied_do_with_repeat_factor
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

program t
    integer :: v(6)
    data (v(i), i = 1, 6) /3*1, 3*9/
    print *, v(1), v(4)
end program t
