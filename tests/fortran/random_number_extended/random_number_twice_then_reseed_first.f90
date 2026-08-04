! vybe-test: fortran/random_number_extended/random_number_twice_then_reseed_first
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
integer :: seed(1) = [42]
real :: r1, r2, r3
call random_seed(put=seed)
call random_number(r1)
call random_number(r2)
call random_seed(put=seed)
call random_number(r3)
if ((merge(1, 0, r3 == r1)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, r3 == r1), "]"
    stop 1
end if
end program t
