! vybe-test: fortran/result_variables/result_variables_07
! origin: languages/fortran/tests/fortran/test_result_variables.rs
program t
if ((f(3)) /= 3) then
    print *, "FAIL: want [3] got [", (f(3)), "]"
    stop 1
end if
contains
integer function f(n) result(r)
integer::n
r=n
end function f
end program t
