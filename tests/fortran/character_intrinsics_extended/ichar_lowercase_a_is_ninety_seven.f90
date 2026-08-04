! vybe-test: fortran/character_intrinsics_extended/ichar_lowercase_a_is_ninety_seven
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
if ((ichar('a')) /= 97) then
    print *, "FAIL: want [97] got [", ichar('a'), "]"
    stop 1
end if
end program t
