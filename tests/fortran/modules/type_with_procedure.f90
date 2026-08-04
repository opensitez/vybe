! vybe-test: fortran/modules/type_with_procedure
! origin: languages/fortran/tests/fortran/test_modules.rs
program t
type :: Counter
integer :: val = 0
contains
procedure :: inc
end type Counter
type(Counter) :: c
print *, c%val
contains
subroutine inc(self)
class(Counter), intent(inout) :: self
self%val = self%val + 1
end subroutine inc
end program t
