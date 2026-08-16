! vybe-test: fortran/derived_types/dt_type_ext_01
! origin: languages/fortran/tests/fortran/test_derived_types.rs
program t
type::b
integer::x
end type b
type,extends(b)::c
integer::y
end type c
type(c) :: obj
obj%x = 4
obj%y = 9
if (obj%x + obj%y /= 13) then
    print *, "FAIL: want [13] got [", obj%x + obj%y, "]"
    stop 1
end if
end program t
