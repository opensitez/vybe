! vybe-test: fortran/expressions/expr_merge_28
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
integer :: x
x = merge(1,2,.true.)
if ((x) /= 1) then
    print *, "FAIL: want [1] got [", x, "]"
    stop 1
end if
end program p
