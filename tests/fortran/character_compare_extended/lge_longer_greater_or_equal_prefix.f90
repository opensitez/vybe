! vybe-test: fortran/character_compare_extended/lge_longer_greater_or_equal_prefix
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((lge('abc', 'ab')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", lge('abc', 'ab'), "]"
    stop 1
end if
end program t
