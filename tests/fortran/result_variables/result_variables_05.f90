! vybe-test: fortran/result_variables/result_variables_05
! origin: languages/fortran/tests/fortran/test_result_variables.rs
logical function f() result(r)
r=.true.
end function f
