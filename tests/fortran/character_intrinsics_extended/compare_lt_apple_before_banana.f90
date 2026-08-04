! vybe-test: fortran/character_intrinsics_extended/compare_lt_apple_before_banana
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
if (('apple' < 'banana') .neqv. .true.) then
    print *, "FAIL: want [true] got [", 'apple' < 'banana', "]"
    stop 1
end if
end program t
