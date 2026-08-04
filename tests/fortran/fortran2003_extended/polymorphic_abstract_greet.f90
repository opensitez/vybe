! vybe-test: fortran/fortran2003_extended/polymorphic_abstract_greet
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
module greetmod
implicit none
type, abstract :: Greeter
contains
procedure(msg_iface), deferred :: message
end type Greeter
abstract interface
function msg_iface(self) result(s)
import Greeter
class(Greeter), intent(in) :: self
character(len=16) :: s
end function msg_iface
end interface
type, extends(Greeter) :: Hello
contains
procedure :: message => hello_msg
end type Hello
contains
function hello_msg(self) result(s)
class(Hello), intent(in) :: self
character(len=16) :: s
s = 'hello'
end function hello_msg
end module greetmod
program t
use greetmod
type(Hello) :: h
if (trim(trim(h%message())) /= "hello") then
    print *, "FAIL: want [hello] got [", trim(h%message()), "]"
    stop 1
end if
end program t
