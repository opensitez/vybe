! vybe-test: fortran/character_intrinsics_extended/concat_three_segments_with_spaces
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=3) :: p = 'one'
character(len=3) :: q = 'two'
character(len=3) :: r = 'six'
if (trim(trim(p) // ' ' // trim(q) // ' ' // trim(r)) /= "one two six") then
    print *, "FAIL: want [one two six] got [", trim(p) // ' ' // trim(q) // ' ' // trim(r), "]"
    stop 1
end if
end program t
