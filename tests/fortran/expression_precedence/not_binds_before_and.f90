! vybe-test: fortran/expression_precedence/not_binds_before_and
! origin: languages/fortran/tests/fortran/test_expression_precedence.rs
program t
if ((.not. .false. .and. .false.) .neqv. .false.) then
    print *, "FAIL: want [false] got [", .not. .false. .and. .false., "]"
    stop 1
end if
end program t
