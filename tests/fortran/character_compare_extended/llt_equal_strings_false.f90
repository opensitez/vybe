! vybe-test: fortran/character_compare_extended/llt_equal_strings_false
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((llt('same', 'same')) .neqv. .false.) then
    print *, "FAIL: want [false] got [", llt('same', 'same'), "]"
    stop 1
end if
end program t
