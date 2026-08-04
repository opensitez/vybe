! vybe-test: fortran/enum_type_extended/enum_auto_four_items
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: A, B, C, D
end enum
if ((A) /= 0) then
    print *, "FAIL: want [0] got [", A, "]"
    stop 1
end if
if ((D) /= 3) then
    print *, "FAIL: want [3] got [", D, "]"
    stop 1
end if
end program t
