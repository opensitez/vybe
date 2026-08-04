! vybe-test: fortran/character_compare_extended/lex_case_insensitive_simulation
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
character(len=3) :: a = 'AbC'
character(len=3) :: b = 'aBc'
if ((llt(a, b)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", llt(a, b), "]"
    stop 1
end if
end program t
