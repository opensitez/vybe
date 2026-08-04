! vybe-test: fortran/enum_type_extended/enum_ieor_toggle
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: A = 5, B = 3
end enum
if ((ieor(A, B)) /= 6) then
    print *, "FAIL: want [6] got [", ieor(A, B), "]"
    stop 1
end if
end program t
