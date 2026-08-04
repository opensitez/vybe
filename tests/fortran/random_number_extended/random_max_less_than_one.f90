! vybe-test: fortran/random_number_extended/random_max_less_than_one
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
real :: r(8)
call random_number(r)
if ((merge(1, 0, maxval(r) < 1.0)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, maxval(r) < 1.0), "]"
    stop 1
end if
end program t
