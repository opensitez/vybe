! vybe-test: fortran/enum_type_extended/enum_module_two_sets
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
module dirs
enum, bind(c)
enumerator :: N = 0, S = 1
end enum
enum, bind(c)
enumerator :: E = 0, W = 1
end enum
end module dirs
program t
use dirs
if ((N + E) /= 0) then
    print *, "FAIL: want [0] got [", N + E, "]"
    stop 1
end if
end program t
