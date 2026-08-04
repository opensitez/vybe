! vybe-test: fortran/expressions/expr_ne_13
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
logical :: x
x = 1 /= 2
if ((x) .neqv. .true.) then
    print *, "FAIL: want [true] got [", x, "]"
    stop 1
end if
end program p
