! vybe-test: fortran/expression_precedence/and_binds_tighter_than_or_false_and_true_or
! origin: languages/fortran/tests/fortran/test_expression_precedence.rs
program t
if ((.false. .and. .true. .or. .true.) .neqv. .true.) then
    print *, "FAIL: want [true] got [", .false. .and. .true. .or. .true., "]"
    stop 1
end if
end program t
