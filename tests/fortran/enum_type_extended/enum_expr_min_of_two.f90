! vybe-test: fortran/enum_type_extended/enum_expr_min_of_two
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: LOW = 2, HIGH = 8
end enum
if ((min(LOW, HIGH)) /= 2) then
    print *, "FAIL: want [2] got [", min(LOW, HIGH), "]"
    stop 1
end if
end program t
