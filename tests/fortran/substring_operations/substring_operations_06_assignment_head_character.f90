! vybe-test: fortran/substring_operations/substring_operations_06_assignment_head_character
! origin: languages/fortran/tests/fortran/test_substring_operations.rs
program p
character(len=6) :: s='abcdef'
s(1:1)='X'
if (trim(trim(s)) /= "Xbcdef") then
    print *, "FAIL: want [Xbcdef] got [", trim(s), "]"
    stop 1
end if
end program p
