! vybe-test: fortran/expressions/expr_char_rel_19
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
logical :: x
x = 'a' < 'b'
if ((x) .neqv. .true.) then
    print *, "FAIL: want [true] got [", x, "]"
    stop 1
end if
end program p
