! vybe-test: fortran/enum_type_extended/enum_expr_addition
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: A = 1, B = 2, C = 3
end enum
if ((A + B) /= 3) then
    print *, "FAIL: want [3] got [", A + B, "]"
    stop 1
end if
end program t
