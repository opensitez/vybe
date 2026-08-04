! vybe-test: fortran/random_number_extended/random_count_below_half
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
real :: r(30)
integer :: c
call random_number(r)
c = count(r < 0.5)
if ((merge(1, 0, c >= 0 .and. c <= 30)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, c >= 0 .and. c <= 30), "]"
    stop 1
end if
end program t
