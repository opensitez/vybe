! vybe-test: fortran/statement_free_form_legacy/statement_free_form_legacy_associated_goto_and_arithmetic_if
! origin: languages/fortran/tests/fortran/test_statement_free_form_legacy.rs

  program statement_free_form_legacy_associated_goto_and_arithmetic_if
    integer :: x
    integer, save :: seen = 0
    x = -1
    if (x) 10, 20, 30
10  seen = seen + 1
20  seen = seen + 2
30  seen = seen + 3
    if ((seen) /= 1) then
    print *, "FAIL: want [1] got [", seen, "]"
    stop 1
end if
end program statement_free_form_legacy_associated_goto_and_arithmetic_if
