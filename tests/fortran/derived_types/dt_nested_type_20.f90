! vybe-test: fortran/derived_types/dt_nested_type_20
! origin: languages/fortran/tests/fortran/test_derived_types.rs
program t
type::t1
integer::x
end type t1
type::t2
type(t1)::a
end type t2
type(t2) :: obj
obj%a%x = 6
if (obj%a%x /= 6) then
    print *, "FAIL: want [6] got [", obj%a%x, "]"
    stop 1
end if
end program t
