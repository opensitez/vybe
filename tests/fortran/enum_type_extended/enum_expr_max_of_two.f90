! vybe-test: fortran/enum_type_extended/enum_expr_max_of_two
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: LOW = 2, HIGH = 8
end enum
if ((max(LOW, HIGH)) /= 8) then
    print *, "FAIL: want [8] got [", max(LOW, HIGH), "]"
    stop 1
end if
end program t
