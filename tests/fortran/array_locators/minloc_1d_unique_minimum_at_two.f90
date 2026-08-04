! vybe-test: fortran/array_locators/minloc_1d_unique_minimum_at_two
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(5) = [3, 1, 9, 1, 5]
integer :: loc(1)
loc = minloc(a)
if ((loc(1)) /= 2) then
    print *, "FAIL: want [2] got [", loc(1), "]"
    stop 1
end if
end program t
