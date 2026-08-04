! vybe-test: fortran/character_compare_extended/lgt_letter_after_digit
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((lgt('a', '1')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", lgt('a', '1'), "]"
    stop 1
end if
end program t
