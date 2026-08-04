! vybe-test: fortran/arithmetic_extended/cmp_ne_equal_false_prints_zero
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
integer :: r
if (8 /= 8) then
r = 1
else
r = 0
end if
if ((r) /= 0) then
    print *, "FAIL: want [0] got [", r, "]"
    stop 1
end if
end program t
