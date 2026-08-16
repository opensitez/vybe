! vybe-test: fortran/result_variables/result_variables_05
! origin: languages/fortran/tests/fortran/test_result_variables.rs
program t
if (merge(1, 0, f()) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, f()), "]"
    stop 1
end if
contains
logical function f() result(r)
r=.true.
end function f
end program t
