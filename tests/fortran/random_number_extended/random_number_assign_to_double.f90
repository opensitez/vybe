! vybe-test: fortran/random_number_extended/random_number_assign_to_double
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
double precision :: r
call random_number(r)
if ((merge(1, 0, r >= 0.0d0 .and. r < 1.0d0)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, r >= 0.0d0 .and. r < 1.0d0), "]"
    stop 1
end if
end program t
