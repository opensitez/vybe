! vybe-test: fortran/character_intrinsics_extended/len_declared_ten_len_trim_two
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=10) :: s = 'go'
if ((len(s)) /= 10) then
    print *, "FAIL: want [10] got [", len(s), "]"
    stop 1
end if
if ((len_trim(s)) /= 2) then
    print *, "FAIL: want [2] got [", len_trim(s), "]"
    stop 1
end if
end program t
