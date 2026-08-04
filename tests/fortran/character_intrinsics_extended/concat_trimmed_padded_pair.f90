! vybe-test: fortran/character_intrinsics_extended/concat_trimmed_padded_pair
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=5) :: a = 'Hi   '
character(len=5) :: b = '  Ho'
if (trim(trim(a // b)) /= "Hi     Ho") then
    print *, "FAIL: want [Hi     Ho] got [", trim(a // b), "]"
    stop 1
end if
end program t
