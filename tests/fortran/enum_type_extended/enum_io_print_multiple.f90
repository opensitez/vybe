! vybe-test: fortran/enum_type_extended/enum_io_print_multiple
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: A = 1, B = 2, C = 3
end enum
if ((A) /= 1) then
    print *, "FAIL: want [1] got [", A, "]"
    stop 1
end if
if ((B) /= 2) then
    print *, "FAIL: want [2] got [", B, "]"
    stop 1
end if
if ((C) /= 3) then
    print *, "FAIL: want [3] got [", C, "]"
    stop 1
end if
end program t
