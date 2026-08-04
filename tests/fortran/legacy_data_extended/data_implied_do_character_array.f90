! vybe-test: fortran/legacy_data_extended/data_implied_do_character_array
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

program t
    character(len=4) :: tags(2)
    data (tags(i), i = 1, 2) /'ab  ', 'cd  '/
    print *, tags(1)
end program t
