! vybe-test: fortran/enum_type_extended/enum_bind_c_large_gap
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: A = 100, B = 200, C = 300
end enum
integer :: v = B
if ((v) /= 200) then
    print *, "FAIL: want [200] got [", v, "]"
    stop 1
end if
end program t
