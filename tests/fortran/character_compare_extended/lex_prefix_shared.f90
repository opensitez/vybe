! vybe-test: fortran/character_compare_extended/lex_prefix_shared
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((lge('prefix', 'pre')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", lge('prefix', 'pre'), "]"
    stop 1
end if
end program t
