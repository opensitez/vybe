! vybe-test: fortran/result_variables/result_variables_08
! origin: languages/fortran/tests/fortran/test_result_variables.rs
program t
if (nint((f(1.5)) * 100) /= 150) then
    print *, "FAIL: want [150] got [", nint((f(1.5)) * 100), "]"
    stop 1
end if
contains
real function f(x) result(r)
real::x
r=x
end function f
end program t
