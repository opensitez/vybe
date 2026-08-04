! vybe-test: fortran/character_intrinsics_extended/compare_lt_chain_three_strings_ordered
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
if (('ant' < 'bee') .neqv. .true.) then
    print *, "FAIL: want [true] got [", 'ant' < 'bee', "]"
    stop 1
end if
if (('bee' < 'cow') .neqv. .true.) then
    print *, "FAIL: want [true] got [", 'bee' < 'cow', "]"
    stop 1
end if
if (('ant' < 'cow') .neqv. .true.) then
    print *, "FAIL: want [true] got [", 'ant' < 'cow', "]"
    stop 1
end if
end program t
