! vybe-test: fortran/expressions/expr_logical_group_37
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
logical :: x
x = .true. .and. (.false. .or. .true.)
if ((x) .neqv. .true.) then
    print *, "FAIL: want [true] got [", x, "]"
    stop 1
end if
end program p
