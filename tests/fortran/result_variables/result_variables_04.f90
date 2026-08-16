! vybe-test: fortran/result_variables/result_variables_04
! origin: languages/fortran/tests/fortran/test_result_variables.rs
program t
if (trim(f()) /= "abc") then
    print *, "FAIL: want [abc] got [", trim(f()), "]"
    stop 1
end if
contains
character(len=3) function f() result(r)
r='abc'
end function f
end program t
