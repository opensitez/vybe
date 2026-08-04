! vybe-test: fortran/substring_operations/substring_operations_05_assignment_mid_section
! origin: languages/fortran/tests/fortran/test_substring_operations.rs
program p
character(len=6) :: s='abcdef'
s(2:3)='ZZ'
if (trim(trim(s)) /= "aZZdef") then
    print *, "FAIL: want [aZZdef] got [", trim(s), "]"
    stop 1
end if
end program p
