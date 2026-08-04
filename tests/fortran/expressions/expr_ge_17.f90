! vybe-test: fortran/expressions/expr_ge_17
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
logical :: x
x = 2 >= 1
if ((x) .neqv. .true.) then
    print *, "FAIL: want [true] got [", x, "]"
    stop 1
end if
end program p
