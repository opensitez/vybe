! vybe-test: fortran/derived_types/dt_poly_comp_10
! origin: languages/fortran/tests/fortran/test_derived_types.rs
program driver
type::t
class(*), allocatable :: item
end type t
type(t) :: obj
real :: got
allocate(obj%item, source=2.5)
got = 0.0
select type (payload => obj%item)
type is (real)
    got = payload
end select
if (nint(got * 2) /= 5) then
    print *, "FAIL: want [5] got [", nint(got * 2), "]"
    stop 1
end if
end program driver
