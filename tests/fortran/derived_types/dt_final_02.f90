! vybe-test: fortran/derived_types/dt_final_02
! origin: languages/fortran/tests/fortran/test_derived_types.rs
type::t
integer::x
contains
final::fin
end type t
contains
subroutine fin(x)
type(t)::x
end subroutine fin
