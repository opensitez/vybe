! vybe-test: fortran/oop/oop_abstract_15
! origin: languages/fortran/tests/fortran/test_oop.rs
module m
type,abstract::t
contains
procedure(p),deferred::run
end type t
abstract interface
subroutine p(this)
import t
class(t)::this
end
end interface
type, extends(t) :: runner
integer :: ran = 0
contains
procedure :: run => runner_run
end type runner
contains
subroutine runner_run(this)
class(runner)::this
this%ran = this%ran + 1
end subroutine runner_run
end module m
program driver
use m
type(runner) :: obj
call obj%run()
call obj%run()
if (obj%ran /= 2) then
    print *, "FAIL: want [2] got [", obj%ran, "]"
    stop 1
end if
end program driver
