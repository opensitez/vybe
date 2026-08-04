! vybe-test: fortran/expression_precedence/comparison_before_logical_and
! origin: languages/fortran/tests/fortran/test_expression_precedence.rs
program t
if ((1 + 2 == 3 .and. .true.) .neqv. .true.) then
    print *, "FAIL: want [true] got [", 1 + 2 == 3 .and. .true., "]"
    stop 1
end if
end program t
