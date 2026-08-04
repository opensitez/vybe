! vybe-test: fortran/expressions/expr_concat_18
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
character(len=2) :: s
s = 'a'//'b'
if (trim(s) /= "ab") then
    print *, "FAIL: want [ab] got [", s, "]"
    stop 1
end if
end program p
