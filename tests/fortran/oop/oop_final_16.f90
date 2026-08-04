! vybe-test: fortran/oop/oop_final_16
! origin: languages/fortran/tests/fortran/test_oop.rs
type::t
contains
final::fin
end type t
contains
subroutine fin(x)
type(t)::x
end subroutine fin
