! vybe-test: fortran/statement_free_form_legacy/statement_free_form_legacy_legacy_comment_in_column
! origin: languages/fortran/tests/fortran/test_statement_free_form_legacy.rs

program statement_free_form_legacy_legacy_comment_in_column
    integer :: a
    a = 10
    !c   legacy comment form preserved
    if ((a) /= 10) then
    print *, "FAIL: want [10] got [", a, "]"
    stop 1
end if
end program statement_free_form_legacy_legacy_comment_in_column
