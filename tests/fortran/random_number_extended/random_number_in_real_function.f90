! vybe-test: fortran/random_number_extended/random_number_in_real_function
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
if ((merge(1, 0, draw() >= 0.0 .and. draw() < 1.0)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, draw() >= 0.0 .and. draw() < 1.0), "]"
    stop 1
end if
contains
function draw() result(r)
real :: r
call random_number(r)
end function draw
end program t
