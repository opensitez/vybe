! vybe-test: fortran/oop/oop_bound_generic_17
! origin: languages/fortran/tests/fortran/test_oop.rs
module m
integer :: hits = 0
type::t
contains
generic::g=>s
procedure::s
end type t
contains
subroutine s(this)
class(t)::this
hits = hits + 1
end subroutine s
end module m
program driver
use m
type(t) :: obj
call obj%s()
if (hits /= 1) then
    print *, "FAIL: want [1] got [", hits, "]"
    stop 1
end if
end program driver
