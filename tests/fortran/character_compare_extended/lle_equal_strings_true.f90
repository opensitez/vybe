! vybe-test: fortran/character_compare_extended/lle_equal_strings_true
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((lle('pair', 'pair')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", lle('pair', 'pair'), "]"
    stop 1
end if
end program t
