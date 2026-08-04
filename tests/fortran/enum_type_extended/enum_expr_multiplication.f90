! vybe-test: fortran/enum_type_extended/enum_expr_multiplication
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: A = 4, B = 5
end enum
if ((A * B) /= 20) then
    print *, "FAIL: want [20] got [", A * B, "]"
    stop 1
end if
end program t
