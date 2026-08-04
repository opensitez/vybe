! vybe-test: fortran/character_compare_extended/lle_space_before_letter
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((lle(' ', 'a')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", lle(' ', 'a'), "]"
    stop 1
end if
end program t
