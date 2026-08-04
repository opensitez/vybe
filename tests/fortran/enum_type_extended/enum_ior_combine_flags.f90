! vybe-test: fortran/enum_type_extended/enum_ior_combine_flags
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: F1 = 1, F2 = 2, F4 = 4
end enum
if ((ior(F1, F2)) /= 3) then
    print *, "FAIL: want [3] got [", ior(F1, F2), "]"
    stop 1
end if
end program t
