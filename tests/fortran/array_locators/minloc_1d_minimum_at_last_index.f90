! vybe-test: fortran/array_locators/minloc_1d_minimum_at_last_index
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(5) = [5, 4, 3, 2, 1]
integer :: loc(1)
loc = minloc(a)
if ((loc(1)) /= 5) then
    print *, "FAIL: want [5] got [", loc(1), "]"
    stop 1
end if
end program t
