! vybe-test: fortran/strings_extended/to_upper_and_to_lower_runtime
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
character(len=5) :: s = 'AbCdE'
if (trim(to_upper(s)) /= "ABCDE") then
    print *, "FAIL: want [ABCDE] got [", to_upper(s), "]"
    stop 1
end if
if (trim(to_lower(s)) /= "abcde") then
    print *, "FAIL: want [abcde] got [", to_lower(s), "]"
    stop 1
end if
end program t
