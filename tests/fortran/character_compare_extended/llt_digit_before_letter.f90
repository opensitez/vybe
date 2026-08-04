! vybe-test: fortran/character_compare_extended/llt_digit_before_letter
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((llt('1', 'a')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", llt('1', 'a'), "]"
    stop 1
end if
end program t
