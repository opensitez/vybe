! vybe-test: fortran/enum_type_extended/enum_io_print_expression
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: BASE = 10, OFFSET = 3
end enum
if ((BASE + OFFSET) /= 13) then
    print *, "FAIL: want [13] got [", BASE + OFFSET, "]"
    stop 1
end if
end program t
