! vybe-test: fortran/array_locators/minloc_1d_ascending_first_is_min
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(5) = [1, 2, 3, 4, 5]
integer :: loc(1)
loc = minloc(a)
if ((loc(1)) /= 1) then
    print *, "FAIL: want [1] got [", loc(1), "]"
    stop 1
end if
end program t
