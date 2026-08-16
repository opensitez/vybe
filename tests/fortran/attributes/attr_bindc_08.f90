! vybe-test: fortran/attributes/attr_bindc_08
! origin: languages/fortran/tests/fortran/test_attributes.rs
program driver
type, bind(c) :: t
integer :: x
end type t
type(t) :: obj
obj%x = 3
if (obj%x /= 3) then
    print *, "FAIL: want [3] got [", obj%x, "]"
    stop 1
end if
end program driver
