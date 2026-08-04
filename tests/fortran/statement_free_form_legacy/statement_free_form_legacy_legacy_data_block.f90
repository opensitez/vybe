! vybe-test: fortran/statement_free_form_legacy/statement_free_form_legacy_legacy_data_block
! origin: languages/fortran/tests/fortran/test_statement_free_form_legacy.rs

program statement_free_form_legacy_legacy_data_block
    integer :: x, y
    data x /3/ y /4/
    if ((x + y) /= 7) then
    print *, "FAIL: want [7] got [", x + y, "]"
    stop 1
end if
end program statement_free_form_legacy_legacy_data_block
