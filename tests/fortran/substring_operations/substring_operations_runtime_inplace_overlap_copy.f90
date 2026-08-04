! vybe-test: fortran/substring_operations/substring_operations_runtime_inplace_overlap_copy
! origin: languages/fortran/tests/fortran/test_substring_operations.rs

program p
character(len=8) :: s
s = 'abcdefgh'
s(3:6) = s(4:7)
if (trim(trim(s)) /= "abdefghh") then
    print *, "FAIL: want [abdefghh] got [", trim(s), "]"
    stop 1
end if
end program p
