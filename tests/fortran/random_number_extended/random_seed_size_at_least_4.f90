! vybe-test: fortran/random_number_extended/random_seed_size_at_least_4
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
integer :: sz
call random_seed(size=sz)
if ((merge(1, 0, sz >= 4)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, sz >= 4), "]"
    stop 1
end if
end program t
