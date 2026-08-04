! vybe-test: fortran/character_intrinsics_extended/iachar_achar_roundtrip_lowercase
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
if (trim(achar(iachar('z'))) /= "z") then
    print *, "FAIL: want [z] got [", achar(iachar('z')), "]"
    stop 1
end if
end program t
