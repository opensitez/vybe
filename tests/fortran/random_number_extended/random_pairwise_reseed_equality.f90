! vybe-test: fortran/random_number_extended/random_pairwise_reseed_equality
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
integer :: seed(1) = [2024]
real :: r1, r2
call random_seed(put=seed)
call random_number(r1)
call random_seed(put=seed)
call random_number(r2)
if ((merge(1, 0, abs(r1 - r2) == 0.0)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, abs(r1 - r2) == 0.0), "]"
    stop 1
end if
end program t
