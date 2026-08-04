! vybe-test: fortran/character_intrinsics_extended/concat_adjusted_strings_with_space
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=8) :: a = '  left'
character(len=8) :: b = 'right  '
if (trim(trim(adjustl(a)) // ' ' // trim(adjustr(b))) /= "left right") then
    print *, "FAIL: want [left right] got [", trim(adjustl(a)) // ' ' // trim(adjustr(b)), "]"
    stop 1
end if
end program t
