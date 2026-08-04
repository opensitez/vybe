! vybe-test: fortran/enum_type_extended/enum_bind_c_cardinal_directions
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: NORTH, SOUTH, EAST, WEST
end enum
integer :: d = EAST
if ((d) /= 2) then
    print *, "FAIL: want [2] got [", d, "]"
    stop 1
end if
end program t
