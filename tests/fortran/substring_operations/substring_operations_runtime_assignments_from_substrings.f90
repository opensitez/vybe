! vybe-test: fortran/substring_operations/substring_operations_runtime_assignments_from_substrings
! origin: languages/fortran/tests/fortran/test_substring_operations.rs

program p
character(len=7) :: s='abcdef'
s(3:5) = s(1:3)
if (trim(trim(s)) /= "ababcef") then
    print *, "FAIL: want [ababcef] got [", trim(s), "]"
    stop 1
end if
end program p
