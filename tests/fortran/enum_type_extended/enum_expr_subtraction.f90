! vybe-test: fortran/enum_type_extended/enum_expr_subtraction
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: A = 10, B = 3
end enum
if ((A - B) /= 7) then
    print *, "FAIL: want [7] got [", A - B, "]"
    stop 1
end if
end program t
