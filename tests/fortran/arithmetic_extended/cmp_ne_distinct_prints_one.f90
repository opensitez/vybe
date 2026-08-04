! vybe-test: fortran/arithmetic_extended/cmp_ne_distinct_prints_one
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
integer :: r
if (8 /= 5) then
r = 1
else
r = 0
end if
if ((r) /= 1) then
    print *, "FAIL: want [1] got [", r, "]"
    stop 1
end if
end program t
