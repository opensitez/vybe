! vybe-test: fortran/array_locators/maxloc_value_at_seven_array
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(5) = [2, 7, 1, 7, 3]
integer :: loc(1)
loc = maxloc(a)
if ((a(loc(1))) /= 7) then
    print *, "FAIL: want [7] got [", a(loc(1)), "]"
    stop 1
end if
end program t
