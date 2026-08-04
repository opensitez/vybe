! vybe-test: fortran/expressions/expr_kind_conv_26
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
integer :: i
real :: r=1.5
i = int(r)
if ((i) /= 1) then
    print *, "FAIL: want [1] got [", i, "]"
    stop 1
end if
end program p
