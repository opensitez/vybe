! vybe-test: fortran/character_intrinsics_extended/char_from_code_seventy_two_is_h
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=1) :: c
c = char(72)
if (trim(c) /= "H") then
    print *, "FAIL: want [H] got [", c, "]"
    stop 1
end if
end program t
