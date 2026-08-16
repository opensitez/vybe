! vybe-test: fortran/derived_types/dt_bindc_08
! origin: languages/fortran/tests/fortran/test_derived_types.rs
program driver
type,bind(c)::t
integer::x
end type t
type(t) :: obj
obj%x = 8
if (obj%x /= 8) then
    print *, "FAIL: want [8] got [", obj%x, "]"
    stop 1
end if
end program driver
