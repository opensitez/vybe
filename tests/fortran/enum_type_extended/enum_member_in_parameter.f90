! vybe-test: fortran/enum_type_extended/enum_member_in_parameter
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: MAX = 100
end enum
integer, parameter :: limit = MAX
if ((limit) /= 100) then
    print *, "FAIL: want [100] got [", limit, "]"
    stop 1
end if
end program t
