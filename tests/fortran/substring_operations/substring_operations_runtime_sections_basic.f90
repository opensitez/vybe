! vybe-test: fortran/substring_operations/substring_operations_runtime_sections_basic
! origin: languages/fortran/tests/fortran/test_substring_operations.rs

program p
character(len=6) :: s='abcdef'
if (trim(trim(s(1:2))) /= "ab") then
    print *, "FAIL: want [ab] got [", trim(s(1:2)), "]"
    stop 1
end if
if (trim(trim(s(2:))) /= "bcdef") then
    print *, "FAIL: want [bcdef] got [", trim(s(2:)), "]"
    stop 1
end if
if (trim(trim(s(:4))) /= "abcd") then
    print *, "FAIL: want [abcd] got [", trim(s(:4)), "]"
    stop 1
end if
if (trim(trim(s(3:3))) /= "c") then
    print *, "FAIL: want [c] got [", trim(s(3:3)), "]"
    stop 1
end if
end program p
