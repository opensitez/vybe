! vybe-test: fortran/expressions/expr_logical_and_09
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
logical :: x
x = .true. .and. .false.
if ((x) .neqv. .false.) then
    print *, "FAIL: want [false] got [", x, "]"
    stop 1
end if
end program p
