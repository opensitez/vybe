! vybe-test: fortran/enum_type_extended/enum_io_print_auto_chain
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: X, Y, Z
end enum
if ((X) /= 0) then
    print *, "FAIL: want [0] got [", X, "]"
    stop 1
end if
if ((Y) /= 1) then
    print *, "FAIL: want [1] got [", Y, "]"
    stop 1
end if
if ((Z) /= 2) then
    print *, "FAIL: want [2] got [", Z, "]"
    stop 1
end if
end program t
