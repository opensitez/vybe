! vybe-test: fortran/abstract_types_deferred/extends_type_of_deferred_child
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
module bases
implicit none
type :: Root
integer :: id = 1
end type Root
type, abstract, extends(Root) :: Node
contains
procedure(val_iface), deferred :: value
end type Node
abstract interface
function val_iface(self) result(v)
import Node
class(Node), intent(in) :: self
integer :: v
end function val_iface
end interface
type, extends(Node) :: Leaf
integer :: payload = 9
contains
procedure :: value => leaf_val
end type Leaf
contains
function leaf_val(self) result(v)
class(Leaf), intent(in) :: self
v = self%payload
end function leaf_val
end module bases
program t
use bases
type(Root) :: r
type(Leaf) :: leaf
if ((extends_type_of(leaf, r)) /= 1) then
    print *, "FAIL: want [1] got [", extends_type_of(leaf, r), "]"
    stop 1
end if
if ((leaf%value()) /= 9) then
    print *, "FAIL: want [9] got [", leaf%value(), "]"
    stop 1
end if
end program t
