! vybe-test: fortran/enum_type_extended/enum_expr_mod
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: A = 17, B = 5
end enum
if ((mod(A, B)) /= 2) then
    print *, "FAIL: want [2] got [", mod(A, B), "]"
    stop 1
end if
end program t
