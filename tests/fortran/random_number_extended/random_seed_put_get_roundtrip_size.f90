! vybe-test: fortran/random_number_extended/random_seed_put_get_roundtrip_size
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
integer :: n, seed(4)
call random_seed(size=n)
call random_seed(put=seed)
call random_seed(get=seed)
if ((merge(1, 0, n == 4)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, n == 4), "]"
    stop 1
end if
end program t
