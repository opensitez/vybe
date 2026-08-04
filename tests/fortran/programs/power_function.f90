! vybe-test: fortran/programs/power_function
! origin: languages/fortran/tests/fortran/test_programs.rs

program test
    if ((2 ** 8) /= 256) then
    print *, "FAIL: want [256] got [", 2 ** 8, "]"
    stop 1
end if
end program test
