! vybe-test: fortran/c_interoperability/c_bind_type_07
! origin: languages/fortran/tests/fortran/test_c_interoperability.rs
program driver
use iso_c_binding
type, bind(c) :: t
 integer(c_int) :: x
end type t
type(t) :: obj
obj%x = 5
if (obj%x /= 5) then
    print *, "FAIL: want [5] got [", obj%x, "]"
    stop 1
end if
end program driver
