! vybe-test: fortran/attributes/attr_extends_10
! origin: languages/fortran/tests/fortran/test_attributes.rs
program t
type :: b
integer::x
end type b
type, extends(b) :: c
integer::y
end type c
type(c) :: obj
obj%x = 2
obj%y = 5
if (obj%x + obj%y /= 7) then
    print *, "FAIL: want [7] got [", obj%x + obj%y, "]"
    stop 1
end if
if (obj%b%x /= 2) then
    print *, "FAIL: want [2] got [", obj%b%x, "]"
    stop 1
end if
end program t
