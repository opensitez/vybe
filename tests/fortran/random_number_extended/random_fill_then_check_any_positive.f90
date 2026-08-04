! vybe-test: fortran/random_number_extended/random_fill_then_check_any_positive
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
real :: r(50)
call random_number(r)
if ((merge(1, 0, any(r > 0.0) .or. all(r == 0.0))) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, any(r > 0.0) .or. all(r == 0.0)), "]"
    stop 1
end if
end program t
