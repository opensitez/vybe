! vybe-test: fortran/random_number_extended/random_number_after_seed_not_constant_zero
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
integer :: seed(1) = [99]
real :: r
call random_seed(put=seed)
call random_number(r)
if ((merge(1, 0, r /= 0.0 .or. r == 0.0)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, r /= 0.0 .or. r == 0.0), "]"
    stop 1
end if
end program t
