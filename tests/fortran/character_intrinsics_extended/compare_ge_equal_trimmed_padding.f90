! vybe-test: fortran/character_intrinsics_extended/compare_ge_equal_trimmed_padding
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=5) :: a = 'same '
character(len=5) :: b = 'same'
if ((trim(a) >= trim(b)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", trim(a) >= trim(b), "]"
    stop 1
end if
end program t
