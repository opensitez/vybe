! vybe-test: fortran/oop/oop_polymorphic_comp_20
! origin: languages/fortran/tests/fortran/test_oop.rs
program t
type::box
class(*), allocatable :: item
end type box
type(box) :: bx
integer :: got
allocate(bx%item, source=41)
got = -1
select type (payload => bx%item)
type is (integer)
    got = payload
end select
if (got /= 41) then
    print *, "FAIL: want [41] got [", got, "]"
    stop 1
end if
end program t
