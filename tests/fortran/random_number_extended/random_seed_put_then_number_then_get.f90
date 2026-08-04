! vybe-test: fortran/random_number_extended/random_seed_put_then_number_then_get
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
integer :: s(1) = [888], g(1)
real :: r
call random_seed(put=s)
call random_number(r)
call random_seed(get=g)
if ((merge(1, 0, g(1) == 888 .and. r >= 0.0)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, g(1) == 888 .and. r >= 0.0), "]"
    stop 1
end if
end program t
