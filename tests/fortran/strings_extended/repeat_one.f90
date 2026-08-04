! vybe-test: fortran/strings_extended/repeat_one
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
if (trim(repeat('x', 1)) /= "x") then
    print *, "FAIL: want [x] got [", repeat('x', 1), "]"
    stop 1
end if
end program t
