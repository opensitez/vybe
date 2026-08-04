! vybe-test: fortran/character_intrinsics_extended/achar_sixty_five_is_capital_a
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=1) :: c
c = achar(65)
if (trim(c) /= "A") then
    print *, "FAIL: want [A] got [", c, "]"
    stop 1
end if
end program t
