! vybe-test: fortran/result_variables/result_variables_09
! origin: languages/fortran/tests/fortran/test_result_variables.rs
recursive integer function f(n) result(r)
integer::n
r=1
end function f
