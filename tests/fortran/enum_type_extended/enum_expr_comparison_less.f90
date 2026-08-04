! vybe-test: fortran/enum_type_extended/enum_expr_comparison_less
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: A = 1, B = 9
end enum
if ((A < B) .neqv. .true.) then
    print *, "FAIL: want [true] got [", A < B, "]"
    stop 1
end if
end program t
