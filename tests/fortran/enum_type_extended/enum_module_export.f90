! vybe-test: fortran/enum_type_extended/enum_module_export
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
module colors
enum, bind(c)
enumerator :: RED = 0, GREEN = 1, BLUE = 2
end enum
end module colors
program t
use colors
if ((GREEN) /= 1) then
    print *, "FAIL: want [1] got [", GREEN, "]"
    stop 1
end if
end program t
