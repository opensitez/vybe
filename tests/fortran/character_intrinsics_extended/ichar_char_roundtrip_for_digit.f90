! vybe-test: fortran/character_intrinsics_extended/ichar_char_roundtrip_for_digit
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
if ((char(ichar('7'))) /= 7) then
    print *, "FAIL: want [7] got [", char(ichar('7')), "]"
    stop 1
end if
end program t
