! vybe-test: fortran/expressions/expr_logical_or_10
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
logical :: x
x = .true. .or. .false.
if ((x) .neqv. .true.) then
    print *, "FAIL: want [true] got [", x, "]"
    stop 1
end if
end program p
