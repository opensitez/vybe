! vybe-test: fortran/enumerations/enum_named_values_20
! origin: languages/fortran/tests/fortran/test_enumerations.rs
program t
enum, bind(c)
enumerator :: sunday=0, monday=1
end enum
if (sunday /= 0) then
    print *, "FAIL: want [0] got [", sunday, "]"
    stop 1
end if
if (monday /= 1) then
    print *, "FAIL: want [1] got [", monday, "]"
    stop 1
end if
end program t
