! vybe-test: fortran/character_compare_extended/lgt_longer_after_prefix
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((lgt('abcd', 'abc')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", lgt('abcd', 'abc'), "]"
    stop 1
end if
end program t
