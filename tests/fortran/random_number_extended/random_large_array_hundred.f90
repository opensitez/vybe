! vybe-test: fortran/random_number_extended/random_large_array_hundred
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
real :: r(100)
call random_number(r)
if ((merge(1, 0, count(r >= 0.0 .and. r < 1.0) == 100)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, count(r >= 0.0 .and. r < 1.0) == 100), "]"
    stop 1
end if
end program t
