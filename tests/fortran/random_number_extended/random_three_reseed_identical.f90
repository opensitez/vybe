! vybe-test: fortran/random_number_extended/random_three_reseed_identical
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
integer :: seed(1) = [31415]
real :: r1, r2, r3
call random_seed(put=seed)
call random_number(r1)
call random_seed(put=seed)
call random_number(r2)
call random_seed(put=seed)
call random_number(r3)
if ((merge(1, 0, r1 == r2 .and. r2 == r3)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, r1 == r2 .and. r2 == r3), "]"
    stop 1
end if
end program t
