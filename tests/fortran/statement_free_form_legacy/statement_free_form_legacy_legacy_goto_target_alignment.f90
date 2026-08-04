! vybe-test: fortran/statement_free_form_legacy/statement_free_form_legacy_legacy_goto_target_alignment
! origin: languages/fortran/tests/fortran/test_statement_free_form_legacy.rs

program statement_free_form_legacy_legacy_goto_target_alignment
    integer :: x
    x = 1
    if (x .eq. 1) goto 5
    x = 0
5   print *, x
end program statement_free_form_legacy_legacy_goto_target_alignment
