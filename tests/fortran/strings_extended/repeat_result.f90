! vybe-test: fortran/strings_extended/repeat_result
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
if (trim(repeat('ab', 3)) /= "ababab") then
    print *, "FAIL: want [ababab] got [", repeat('ab', 3), "]"
    stop 1
end if
end program t
