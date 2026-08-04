! vybe-test: fortran/variables/double_precision_var
! origin: languages/fortran/tests/fortran/test_variables.rs
program t
double precision :: d = 1.23456789d0
if ((nint(d*1_8)) /= 1) then
    print *, "FAIL: want [1] got [", nint(d*1_8), "]"
    stop 1
end if
end program t
