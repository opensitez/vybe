! vybe-test: fortran/specification_part/spec_import_22
! origin: languages/fortran/tests/fortran/test_specification_part.rs
module m
implicit none
type :: payload
 integer :: v = 0
end type
interface
 subroutine consume(p)
  import :: payload
  type(payload), intent(inout) :: p
 end subroutine consume
end interface
end module m
subroutine consume(p)
use m, only: payload
implicit none
type(payload), intent(inout) :: p
p%v = p%v + 5
end subroutine consume
program t
use m
implicit none
type(payload) :: obj
obj%v = 2
call consume(obj)
if (obj%v /= 7) then
    print *, "FAIL: want [7] got [", obj%v, "]"
    stop 1
end if
end program t
