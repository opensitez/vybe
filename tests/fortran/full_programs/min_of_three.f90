! vybe-test: fortran/full_programs/min_of_three
! origin: languages/fortran/tests/fortran/test_full_programs.rs
program t
if ((min(min(5, 3), 7)) /= 3) then
    print *, "FAIL: want [3] got [", min(min(5, 3), 7), "]"
    stop 1
end if
end program t
