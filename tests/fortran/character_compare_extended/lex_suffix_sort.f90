! vybe-test: fortran/character_compare_extended/lex_suffix_sort
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((llt('file.txt', 'file.txz')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", llt('file.txt', 'file.txz'), "]"
    stop 1
end if
end program t
