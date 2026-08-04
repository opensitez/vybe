! vybe-test: fortran/statement_free_form_legacy/statement_free_form_legacy_hollerith_data_literal
! origin: languages/fortran/tests/fortran/test_statement_free_form_legacy.rs

program statement_free_form_legacy_hollerith_data_literal
    character*4 c
    data c /4hABCD/
    if (trim(c) /= "ABCD") then
    print *, "FAIL: want [ABCD] got [", c, "]"
    stop 1
end if
end program statement_free_form_legacy_hollerith_data_literal
