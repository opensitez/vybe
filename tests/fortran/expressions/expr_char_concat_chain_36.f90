! vybe-test: fortran/expressions/expr_char_concat_chain_36
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
character(len=3) :: s
s = 'a'//'b'//'c'
if (trim(s) /= "abc") then
    print *, "FAIL: want [abc] got [", s, "]"
    stop 1
end if
end program p
