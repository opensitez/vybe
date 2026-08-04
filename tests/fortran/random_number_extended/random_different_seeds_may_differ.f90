! vybe-test: fortran/random_number_extended/random_different_seeds_may_differ
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
integer :: s1(1) = [1], s2(1) = [2]
real :: r1, r2
call random_seed(put=s1)
call random_number(r1)
call random_seed(put=s2)
call random_number(r2)
if ((merge(1, 0, r1 /= r2 .or. r1 == r2)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, r1 /= r2 .or. r1 == r2), "]"
    stop 1
end if
end program t
