! vybe-test: fortran/full_programs/absolute_difference
! origin: languages/fortran/tests/fortran/test_full_programs.rs
program t
if ((abs(5 - 12)) /= 7) then
    print *, "FAIL: want [7] got [", abs(5 - 12), "]"
    stop 1
end if
end program t
