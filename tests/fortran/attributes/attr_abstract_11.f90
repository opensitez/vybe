! vybe-test: fortran/attributes/attr_abstract_11
! origin: languages/fortran/tests/fortran/test_attributes.rs
program driver
type, abstract :: t
integer::x
end type t
type, extends(t) :: concrete
integer::y
end type concrete
type(concrete) :: obj
obj%x = 4
obj%y = 6
if (obj%x + obj%y /= 10) then
    print *, "FAIL: want [10] got [", obj%x + obj%y, "]"
    stop 1
end if
end program driver
