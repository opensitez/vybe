! vybe-test: fortran/program_units/program_pass_nopass_26
! origin: languages/fortran/tests/fortran/test_program_units.rs
module m
type :: t
contains
 procedure, pass :: s1
 procedure, nopass :: s2
end type t
contains
subroutine s1(this)
 class(t) :: this
end subroutine s1
subroutine s2()
end subroutine s2
end module m
