! vybe-test: fortran/enum_type_extended/enum_expr_abs_negative_member
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: NEG = -7
end enum
if ((abs(NEG)) /= 7) then
    print *, "FAIL: want [7] got [", abs(NEG), "]"
    stop 1
end if
end program t
