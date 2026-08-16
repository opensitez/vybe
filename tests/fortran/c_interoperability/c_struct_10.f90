! vybe-test: fortran/c_interoperability/c_struct_10
! origin: languages/fortran/tests/fortran/test_c_interoperability.rs
program t
use iso_c_binding
type, bind(c) :: point
 integer(c_int) :: x
 integer(c_int) :: y
end type point
type(point) :: p
p%x = 3
p%y = 4
if (p%x * p%y /= 12) then
    print *, "FAIL: want [12] got [", p%x * p%y, "]"
    stop 1
end if
end program t
