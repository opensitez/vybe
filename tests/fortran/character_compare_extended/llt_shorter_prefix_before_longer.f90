! vybe-test: fortran/character_compare_extended/llt_shorter_prefix_before_longer
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((llt('abc', 'abcd')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", llt('abc', 'abcd'), "]"
    stop 1
end if
end program t
