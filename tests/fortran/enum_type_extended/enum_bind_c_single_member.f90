! vybe-test: fortran/enum_type_extended/enum_bind_c_single_member
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: ONLY = 42
end enum
integer :: v = ONLY
if ((v) /= 42) then
    print *, "FAIL: want [42] got [", v, "]"
    stop 1
end if
end program t
