! vybe-test: fortran/enum_type_extended/enum_auto_after_explicit_start
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: START = 10, NEXT, LAST
end enum
if ((START) /= 10) then
    print *, "FAIL: want [10] got [", START, "]"
    stop 1
end if
if ((NEXT) /= 11) then
    print *, "FAIL: want [11] got [", NEXT, "]"
    stop 1
end if
if ((LAST) /= 12) then
    print *, "FAIL: want [12] got [", LAST, "]"
    stop 1
end if
end program t
