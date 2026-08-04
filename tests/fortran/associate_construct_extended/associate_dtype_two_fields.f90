! vybe-test: fortran/associate_construct_extended/associate_dtype_two_fields
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
type :: Coord
integer :: x, y
end type Coord
type(Coord) :: c
c%x = 3
c%y = 4
associate (px => c%x, py => c%y)
if ((px + py) /= 7) then
    print *, "FAIL: want [7] got [", px + py, "]"
    stop 1
end if
end associate
end program t
