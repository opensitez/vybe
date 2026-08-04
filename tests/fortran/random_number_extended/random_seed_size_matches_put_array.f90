! vybe-test: fortran/random_number_extended/random_seed_size_matches_put_array
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
integer :: n, seed(3) = [5,6,7], got(3)
call random_seed(size=n)
call random_seed(put=seed)
call random_seed(get=got)
if ((merge(1, 0, n == 3 .and. got(2) == 6)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, n == 3 .and. got(2) == 6), "]"
    stop 1
end if
end program t
