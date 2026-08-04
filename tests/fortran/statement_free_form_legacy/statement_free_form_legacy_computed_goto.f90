! vybe-test: fortran/statement_free_form_legacy/statement_free_form_legacy_computed_goto
! origin: languages/fortran/tests/fortran/test_statement_free_form_legacy.rs

program statement_free_form_legacy_computed_goto
    integer :: pick
    pick = 2
    goto (10,20,30), pick
10  print *, 'first'
    stop
20  print *, 'second'
    stop
30  print *, 'third'
    stop
end program statement_free_form_legacy_computed_goto
