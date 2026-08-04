! vybe-test: fortran/character_compare_extended/lge_equal_strings_true
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((lge('pair', 'pair')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", lge('pair', 'pair'), "]"
    stop 1
end if
end program t
