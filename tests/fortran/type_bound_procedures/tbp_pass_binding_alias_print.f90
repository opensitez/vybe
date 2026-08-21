! vybe-test: fortran/type_bound_procedures/tbp_pass_binding_alias_print
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
module m
type :: Label
character(len=8) :: text = 'vybe'
contains
procedure :: show => emit_label
end type Label
contains
subroutine emit_label(self)
class(Label), intent(in) :: self
if (trim(trim(self%text)) /= "vybe") then
    print *, "FAIL: want [vybe] got [", trim(self%text), "]"
    stop 1
end if
end subroutine emit_label
end module m
program driver
use m
type(Label) :: item
call item%show()
end program driver
