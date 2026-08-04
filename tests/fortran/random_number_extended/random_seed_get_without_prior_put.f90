! vybe-test: fortran/random_number_extended/random_seed_get_without_prior_put
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
integer :: seed(8)
call random_seed(get=seed)
if ((merge(1, 0, size(seed) >= 1)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, size(seed) >= 1), "]"
    stop 1
end if
end program t
