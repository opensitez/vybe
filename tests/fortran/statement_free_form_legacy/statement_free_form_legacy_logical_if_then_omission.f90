! vybe-test: fortran/statement_free_form_legacy/statement_free_form_legacy_logical_if_then_omission
! origin: languages/fortran/tests/fortran/test_statement_free_form_legacy.rs

program statement_free_form_legacy_logical_if_then_omission
    logical :: done
    done = .false.
    if (done) print *, 'skip'
    if (.not. done) print *, 'run'
end program statement_free_form_legacy_logical_if_then_omission
