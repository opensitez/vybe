! vybe-test: fortran/enum_type_extended/enum_expr_division
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: A = 20, B = 4
end enum
if ((A / B) /= 5) then
    print *, "FAIL: want [5] got [", A / B, "]"
    stop 1
end if
end program t
