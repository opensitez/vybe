! vybe-test: fortran/fortran2003_extended/class_allocatable_abstract_dispatch
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
module animals
implicit none
type, abstract :: Animal
integer :: legs = 0
contains
procedure(legs_iface), deferred :: count_legs
end type Animal
abstract interface
function legs_iface(self) result(n)
import Animal
class(Animal), intent(in) :: self
integer :: n
end function legs_iface
end interface
type, extends(Animal) :: Spider
contains
procedure :: count_legs => spider_legs
end type Spider
contains
function spider_legs(self) result(n)
class(Spider), intent(in) :: self
n = 8
end function spider_legs
end module animals
program t
use animals
class(Animal), allocatable :: a
allocate(Spider :: a)
if ((a%count_legs()) /= 8) then
    print *, "FAIL: want [8] got [", a%count_legs(), "]"
    stop 1
end if
end program t
