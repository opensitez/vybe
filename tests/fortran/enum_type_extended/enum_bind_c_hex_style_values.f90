! vybe-test: fortran/enum_type_extended/enum_bind_c_hex_style_values
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: FLAG_A = 1, FLAG_B = 2, FLAG_C = 4, FLAG_D = 8
end enum
integer :: f = FLAG_C
if ((f) /= 4) then
    print *, "FAIL: want [4] got [", f, "]"
    stop 1
end if
end program t
