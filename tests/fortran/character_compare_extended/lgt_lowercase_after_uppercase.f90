! vybe-test: fortran/character_compare_extended/lgt_lowercase_after_uppercase
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((lgt('a', 'A')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", lgt('a', 'A'), "]"
    stop 1
end if
end program t
