! vybe-test: fortran/full_programs/triangle_area
! origin: languages/fortran/tests/fortran/test_full_programs.rs
program t
real :: base, height, area
base = 10.0
height = 5.0
area = 0.5 * base * height
if ((area) /= 25) then
    print *, "FAIL: want [25] got [", area, "]"
    stop 1
end if
end program t
