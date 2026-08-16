! vybe-test: fortran/result_variables/result_variables_03
! origin: languages/fortran/tests/fortran/test_result_variables.rs
program t
if (nint(real(f()) + aimag(f())) /= 3) then
    print *, "FAIL: want [3] got [", nint(real(f()) + aimag(f())), "]"
    stop 1
end if
contains
complex function f() result(r)
r=(1.0,2.0)
end function f
end program t
