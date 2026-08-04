! vybe-test: fortran/legacy_data_extended/common_character_pair_len
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs
program t
character(len=3) :: s1, s2
common /txt/ s1, s2
s1 = 'ab'
s2 = 'cd'
if ((len_trim(s1)) /= 2) then
    print *, "FAIL: want [2] got [", len_trim(s1), "]"
    stop 1
end if
if ((len_trim(s2)) /= 2) then
    print *, "FAIL: want [2] got [", len_trim(s2), "]"
    stop 1
end if
end program t
