! vybe-test: fortran/character_compare_extended/lle_prefix_less_or_equal
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((lle('ab', 'abc')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", lle('ab', 'abc'), "]"
    stop 1
end if
end program t
