! vybe-test: fortran/enum_type_extended/enum_two_independent_sets
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: RED = 0, GREEN = 1
end enum
enum, bind(c)
enumerator :: UP = 0, DOWN = 1
end enum
if ((RED + DOWN) /= 1) then
    print *, "FAIL: want [1] got [", RED + DOWN, "]"
    stop 1
end if
end program t
