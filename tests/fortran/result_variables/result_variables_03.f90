! vybe-test: fortran/result_variables/result_variables_03
! origin: languages/fortran/tests/fortran/test_result_variables.rs
complex function f() result(r)
r=(1.0,2.0)
end function f
