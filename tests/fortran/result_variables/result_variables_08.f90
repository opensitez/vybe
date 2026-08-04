! vybe-test: fortran/result_variables/result_variables_08
! origin: languages/fortran/tests/fortran/test_result_variables.rs
real function f(x) result(r)
real::x
r=x
end function f
