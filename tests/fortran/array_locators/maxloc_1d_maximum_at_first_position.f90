! vybe-test: fortran/array_locators/maxloc_1d_maximum_at_first_position
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(4) = [99, 1, 2, 3]
integer :: loc(1)
loc = maxloc(a)
if ((loc(1)) /= 1) then
    print *, "FAIL: want [1] got [", loc(1), "]"
    stop 1
end if
end program t
