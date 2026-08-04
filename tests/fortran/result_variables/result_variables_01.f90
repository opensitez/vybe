! vybe-test: fortran/result_variables/result_variables_01
! origin: languages/fortran/tests/fortran/test_result_variables.rs
integer function f() result(r)
r=1
end function f
