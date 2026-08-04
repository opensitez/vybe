! vybe-test: fortran/character_compare_extended/lex_number_prefix_vs_full
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((llt('12', '123')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", llt('12', '123'), "]"
    stop 1
end if
end program t
