! vybe-test: fortran/character_compare_extended/lgt_equal_strings_false
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((lgt('same', 'same')) .neqv. .false.) then
    print *, "FAIL: want [false] got [", lgt('same', 'same'), "]"
    stop 1
end if
end program t
