! vybe-test: fortran/character_compare_extended/lex_tab_vs_space
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program driver
character(len=1) :: t = char(9)
if ((lgt(t, ' ')) .neqv. .false.) then
    print *, "FAIL: want [false] got [", lgt(t, ' '), "]"
    stop 1
end if
end program driver