! vybe-test: fortran/result_variables/result_variables_04
! origin: languages/fortran/tests/fortran/test_result_variables.rs
character(len=3) function f() result(r)
r='abc'
end function f
