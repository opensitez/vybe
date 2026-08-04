! vybe-test: fortran/random_number_extended/random_seed_get_returns_size
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
integer :: n
call random_seed(size=n)
if ((merge(1, 0, n >= 1)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, n >= 1), "]"
    stop 1
end if
end program t
