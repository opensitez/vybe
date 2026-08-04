! vybe-test: fortran/full_programs/celsius_to_fahrenheit
! origin: languages/fortran/tests/fortran/test_full_programs.rs
program t
real :: c, f
c = 100.0
f = c * 9.0 / 5.0 + 32.0
if ((f) /= 212) then
    print *, "FAIL: want [212] got [", f, "]"
    stop 1
end if
end program t
