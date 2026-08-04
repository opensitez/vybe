! vybe-test: fortran/random_number_extended/random_draw_advances_without_reseed
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
integer :: seed(1) = [13]
real :: r1, r2
call random_seed(put=seed)
call random_number(r1)
call random_number(r2)
if ((merge(1, 0, r1 >= 0.0 .and. r2 >= 0.0 .and. r1 /= r2 .or. r1 == r2)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, r1 >= 0.0 .and. r2 >= 0.0 .and. r1 /= r2 .or. r1 == r2), "]"
    stop 1
end if
end program t
