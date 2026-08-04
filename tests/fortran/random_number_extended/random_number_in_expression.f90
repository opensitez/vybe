! vybe-test: fortran/random_number_extended/random_number_in_expression
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
real :: r, x
call random_number(r)
x = r * 0.0 + 0.5
if ((merge(1, 0, x == 0.5)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, x == 0.5), "]"
    stop 1
end if
end program t
