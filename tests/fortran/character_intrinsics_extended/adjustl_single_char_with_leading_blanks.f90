! vybe-test: fortran/character_intrinsics_extended/adjustl_single_char_with_leading_blanks
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=6) :: s = '    Z'
if (trim(trim(adjustl(s))) /= "Z") then
    print *, "FAIL: want [Z] got [", trim(adjustl(s)), "]"
    stop 1
end if
end program t
