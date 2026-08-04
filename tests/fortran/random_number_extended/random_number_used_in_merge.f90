! vybe-test: fortran/random_number_extended/random_number_used_in_merge
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
real :: r
call random_number(r)
if ((merge(1, 0, r < 1.0)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, r < 1.0), "]"
    stop 1
end if
end program t
