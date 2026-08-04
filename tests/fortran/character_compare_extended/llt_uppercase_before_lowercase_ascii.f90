! vybe-test: fortran/character_compare_extended/llt_uppercase_before_lowercase_ascii
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((llt('A', 'a')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", llt('A', 'a'), "]"
    stop 1
end if
end program t
