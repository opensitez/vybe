! vybe-test: fortran/random_number_extended/random_number_floor_is_zero_or_more
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
real :: r
call random_number(r)
if ((merge(1, 0, int(r) >= 0)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, int(r) >= 0), "]"
    stop 1
end if
end program t
