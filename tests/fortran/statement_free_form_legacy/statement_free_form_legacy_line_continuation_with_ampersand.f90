! vybe-test: fortran/statement_free_form_legacy/statement_free_form_legacy_line_continuation_with_ampersand
! origin: languages/fortran/tests/fortran/test_statement_free_form_legacy.rs

program statement_free_form_legacy_line_continuation_with_ampersand
    integer :: value
    value = 1 + &
            2 + &
            3
    if ((value) /= 6) then
    print *, "FAIL: want [6] got [", value, "]"
    stop 1
end if
end program statement_free_form_legacy_line_continuation_with_ampersand
