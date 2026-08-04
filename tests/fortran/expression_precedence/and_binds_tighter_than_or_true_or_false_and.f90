! vybe-test: fortran/expression_precedence/and_binds_tighter_than_or_true_or_false_and
! origin: languages/fortran/tests/fortran/test_expression_precedence.rs
program t
if ((.true. .or. .false. .and. .false.) .neqv. .true.) then
    print *, "FAIL: want [true] got [", .true. .or. .false. .and. .false., "]"
    stop 1
end if
end program t
