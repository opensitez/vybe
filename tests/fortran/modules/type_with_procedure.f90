! vybe-test: fortran/modules/type_with_procedure
! origin: languages/fortran/tests/fortran/test_modules.rs
module m
type :: Counter
integer :: val = 0
contains
procedure :: inc
end type Counter
contains
subroutine inc(self)
class(Counter), intent(inout) :: self
self%val = self%val + 1
end subroutine inc
end module m
program driver
use m
type(Counter) :: c
print *, c%val
end program driver
