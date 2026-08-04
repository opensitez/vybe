! vybe-test: fortran/enum_type_extended/enum_io_print_single
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: VAL = 42
end enum
if ((VAL) /= 42) then
    print *, "FAIL: want [42] got [", VAL, "]"
    stop 1
end if
end program t
