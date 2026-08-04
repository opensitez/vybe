! vybe-test: fortran/expressions/expr_masked_where_30
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
integer :: a(3)=[1,2,3]
where (a > 1) a = a + 1
if ((sum(a)) /= 8) then
    print *, "FAIL: want [8] got [", sum(a), "]"
    stop 1
end if
end program p
