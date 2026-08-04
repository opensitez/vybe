! vybe-test: fortran/array_locators/minloc_1d_negative_minimum_at_two
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(3) = [-1, -5, -2]
integer :: loc(1)
loc = minloc(a)
if ((loc(1)) /= 2) then
    print *, "FAIL: want [2] got [", loc(1), "]"
    stop 1
end if
end program t
