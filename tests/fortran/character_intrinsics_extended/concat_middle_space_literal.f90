! vybe-test: fortran/character_intrinsics_extended/concat_middle_space_literal
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=4) :: a = 'ab  '
character(len=4) :: b = '  cd'
if (trim(trim(trim(a) // ' ' // trim(b))) /= "ab   cd") then
    print *, "FAIL: want [ab   cd] got [", trim(trim(a) // ' ' // trim(b)), "]"
    stop 1
end if
end program t
