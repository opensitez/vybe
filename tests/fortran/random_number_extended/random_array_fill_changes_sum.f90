! vybe-test: fortran/random_number_extended/random_array_fill_changes_sum
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
real :: a(10)
call random_number(a)
if ((merge(1, 0, sum(a) >= 0.0)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, sum(a) >= 0.0), "]"
    stop 1
end if
end program t
