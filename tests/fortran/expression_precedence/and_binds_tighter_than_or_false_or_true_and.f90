! vybe-test: fortran/expression_precedence/and_binds_tighter_than_or_false_or_true_and
! origin: languages/fortran/tests/fortran/test_expression_precedence.rs
program t
if ((.false. .or. .true. .and. .false.) .neqv. .false.) then
    print *, "FAIL: want [false] got [", .false. .or. .true. .and. .false., "]"
    stop 1
end if
end program t
