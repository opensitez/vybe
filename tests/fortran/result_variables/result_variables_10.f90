! vybe-test: fortran/result_variables/result_variables_10
! origin: languages/fortran/tests/fortran/test_result_variables.rs
program t
if ((f()) /= 1) then
    print *, "FAIL: want [1] got [", (f()), "]"
    stop 1
end if
contains
function f() result(r)
integer :: r
r=1
end function f
end program t
