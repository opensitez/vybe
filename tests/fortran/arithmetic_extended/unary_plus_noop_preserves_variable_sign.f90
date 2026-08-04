! vybe-test: fortran/arithmetic_extended/unary_plus_noop_preserves_variable_sign
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
integer :: x
x = -9
if ((+x) /= -9) then
    print *, "FAIL: want [-9] got [", +x, "]"
    stop 1
end if
end program t
