! vybe-test: fortran/result_variables/result_variables_06
! origin: languages/fortran/tests/fortran/test_result_variables.rs
type t
 integer::x
end type t
type(t) function f() result(r)
r%x=1
end function f
