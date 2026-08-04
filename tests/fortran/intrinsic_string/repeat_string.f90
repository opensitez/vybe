! vybe-test: fortran/intrinsic_string/repeat_string
! origin: languages/fortran/tests/fortran/test_intrinsic_string.rs
program t
if (trim(repeat("ab", 3)) /= "ababab") then
    print *, "FAIL: want [ababab] got [", repeat("ab", 3), "]"
    stop 1
end if
end program t
