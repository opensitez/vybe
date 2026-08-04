! vybe-test: fortran/result_variables/result_variables_07
! origin: languages/fortran/tests/fortran/test_result_variables.rs
integer function f(n) result(r)
integer::n
r=n
end function f
