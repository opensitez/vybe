! vybe-test: fortran/procedure_results/procedure_results_08
! origin: languages/fortran/tests/fortran/test_procedure_results.rs
type t
 integer::x
end type t
type(t) function f()
f%x=1
end function f
