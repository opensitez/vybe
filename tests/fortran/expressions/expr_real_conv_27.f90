! vybe-test: fortran/expressions/expr_real_conv_27
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
real :: r
r = real(3)
if ((r) /= 3) then
    print *, "FAIL: want [3] got [", r, "]"
    stop 1
end if
end program p
