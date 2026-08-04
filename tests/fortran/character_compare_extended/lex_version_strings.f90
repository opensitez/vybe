! vybe-test: fortran/character_compare_extended/lex_version_strings
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((llt('1.09', '1.10')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", llt('1.09', '1.10'), "]"
    stop 1
end if
end program t
