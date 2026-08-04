! vybe-test: fortran/enum_type_extended/enum_bind_c_negative_start
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: MINUS2 = -2, MINUS1, ZERO
end enum
integer :: v = ZERO
if ((v) /= 0) then
    print *, "FAIL: want [0] got [", v, "]"
    stop 1
end if
end program t
