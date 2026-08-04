! vybe-test: fortran/enum_type_extended/enum_auto_mid_chain_restart
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: X = 5, Y, Z = 20, W
end enum
if ((Y) /= 6) then
    print *, "FAIL: want [6] got [", Y, "]"
    stop 1
end if
if ((W) /= 21) then
    print *, "FAIL: want [21] got [", W, "]"
    stop 1
end if
end program t
