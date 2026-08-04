! vybe-test: fortran/random_number_extended/random_array_section_fill
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
real :: a(6)
call random_number(a(2:5))
if ((merge(1, 0, all(a(2:5) >= 0.0 .and. a(2:5) < 1.0))) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, all(a(2:5) >= 0.0 .and. a(2:5) < 1.0)), "]"
    stop 1
end if
end program t
