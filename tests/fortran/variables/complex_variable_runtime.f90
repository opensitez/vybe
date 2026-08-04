! vybe-test: fortran/variables/complex_variable_runtime
! origin: languages/fortran/tests/fortran/test_variables.rs
program t
complex :: c = (1.25, -2.5)
if ((nint(real(c)*10)) /= 12) then
    print *, "FAIL: want [12] got [", nint(real(c)*10), "]"
    stop 1
end if
if ((nint(aimag(c)*10)) /= -25) then
    print *, "FAIL: want [-25] got [", nint(aimag(c)*10), "]"
    stop 1
end if
end program t
