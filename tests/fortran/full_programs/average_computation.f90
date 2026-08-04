! vybe-test: fortran/full_programs/average_computation
! origin: languages/fortran/tests/fortran/test_full_programs.rs
program t
real :: avg
avg = (10 + 20 + 30) / 3.0
if ((avg) /= 20) then
    print *, "FAIL: want [20] got [", avg, "]"
    stop 1
end if
end program t
