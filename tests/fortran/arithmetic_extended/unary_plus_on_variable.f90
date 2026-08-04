! vybe-test: fortran/arithmetic_extended/unary_plus_on_variable
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
integer :: x = 42
if ((+x) /= 42) then
    print *, "FAIL: want [42] got [", +x, "]"
    stop 1
end if
end program t
