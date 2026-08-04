! vybe-test: fortran/enum_type_extended/enum_bind_c_rgb_values
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: RED = 0, GREEN = 1, BLUE = 2
end enum
integer :: c = GREEN
if ((c) /= 1) then
    print *, "FAIL: want [1] got [", c, "]"
    stop 1
end if
end program t
