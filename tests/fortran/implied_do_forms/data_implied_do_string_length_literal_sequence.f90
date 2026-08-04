! vybe-test: fortran/implied_do_forms/data_implied_do_string_length_literal_sequence
! origin: languages/fortran/tests/fortran/test_implied_do_forms.rs

program t
    character(len=1) :: tags(3)
    data (tags(i), i = 1, 3) /'a','b','c'/
    print *, tags(2)
end program t
