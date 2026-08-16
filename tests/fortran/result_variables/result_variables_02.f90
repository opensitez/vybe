! vybe-test: fortran/result_variables/result_variables_02
! origin: languages/fortran/tests/fortran/test_result_variables.rs
program t
if (nint((f()) * 100) /= 100) then
    print *, "FAIL: want [100] got [", nint((f()) * 100), "]"
    stop 1
end if
contains
real function f() result(r)
r=1.0
end function f
end program t
