! vybe-test: fortran/result_variables/result_variables_09
! origin: languages/fortran/tests/fortran/test_result_variables.rs
program t
if ((f(3)) /= 1) then
    print *, "FAIL: want [1] got [", (f(3)), "]"
    stop 1
end if
contains
recursive integer function f(n) result(r)
integer::n
r=1
end function f
end program t
