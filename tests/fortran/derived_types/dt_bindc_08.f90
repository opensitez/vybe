! vybe-test: fortran/derived_types/dt_bindc_08
! origin: languages/fortran/tests/fortran/test_derived_types.rs
type,bind(c)::t
integer::x
end type t
