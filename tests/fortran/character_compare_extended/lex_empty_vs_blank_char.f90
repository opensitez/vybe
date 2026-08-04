! vybe-test: fortran/character_compare_extended/lex_empty_vs_blank_char
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
character(len=1) :: a = ' '
if ((lle(a, ' ')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", lle(a, ' '), "]"
    stop 1
end if
end program t
