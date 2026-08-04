! vybe-test: fortran/legacy_data_extended/data_implied_do_scalar_and_array_mix
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

program t
    integer :: head, tail(2)
    data head /0/, (tail(i), i = 1, 2) /4, 5/
    print *, head + tail(2)
end program t
