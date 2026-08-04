! vybe-test: fortran/statement_free_form_legacy/statement_free_form_legacy_assigned_goto
! origin: languages/fortran/tests/fortran/test_statement_free_form_legacy.rs

    program statement_free_form_legacy_assigned_goto
    integer :: target
    assign 10 to target
    goto target
    if (trim('unexpected') /= "assigned") then
    print *, "FAIL: want [assigned] got [", 'unexpected', "]"
    stop 1
end if
10  print *, 'assigned'
end program statement_free_form_legacy_assigned_goto
