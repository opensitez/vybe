! vybe-test: fortran/character_compare_extended/lge_letter_greater_or_equal_space
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((lge('a', ' ')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", lge('a', ' '), "]"
    stop 1
end if
end program t
