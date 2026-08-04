! vybe-test: fortran/character_intrinsics_extended/adjustr_strips_trailing_blanks_on_padded
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=12) :: s = 'Code      '
if (trim(trim(adjustr(s))) /= "Code") then
    print *, "FAIL: want [Code] got [", trim(adjustr(s)), "]"
    stop 1
end if
end program t
