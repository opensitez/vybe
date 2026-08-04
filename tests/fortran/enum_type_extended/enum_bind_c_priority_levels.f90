! vybe-test: fortran/enum_type_extended/enum_bind_c_priority_levels
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: LOW = 1, MEDIUM = 5, HIGH = 10
end enum
integer :: p = HIGH
if ((p) /= 10) then
    print *, "FAIL: want [10] got [", p, "]"
    stop 1
end if
end program t
