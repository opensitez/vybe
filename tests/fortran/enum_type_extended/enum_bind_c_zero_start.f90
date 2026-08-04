! vybe-test: fortran/enum_type_extended/enum_bind_c_zero_start
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: FIRST = 0, SECOND, THIRD
end enum
integer :: v = THIRD
if ((v) /= 2) then
    print *, "FAIL: want [2] got [", v, "]"
    stop 1
end if
end program t
