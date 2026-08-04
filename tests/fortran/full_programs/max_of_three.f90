! vybe-test: fortran/full_programs/max_of_three
! origin: languages/fortran/tests/fortran/test_full_programs.rs
program t
if ((max(max(5, 3), 7)) /= 7) then
    print *, "FAIL: want [7] got [", max(max(5, 3), 7), "]"
    stop 1
end if
end program t
