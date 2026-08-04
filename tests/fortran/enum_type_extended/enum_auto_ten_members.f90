! vybe-test: fortran/enum_type_extended/enum_auto_ten_members
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: E0, E1, E2, E3, E4, E5, E6, E7, E8, E9
end enum
if ((E0) /= 0) then
    print *, "FAIL: want [0] got [", E0, "]"
    stop 1
end if
if ((E9) /= 9) then
    print *, "FAIL: want [9] got [", E9, "]"
    stop 1
end if
end program t
