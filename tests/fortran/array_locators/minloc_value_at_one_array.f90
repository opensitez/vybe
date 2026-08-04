! vybe-test: fortran/array_locators/minloc_value_at_one_array
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(5) = [2, 7, 1, 7, 3]
integer :: loc(1)
loc = minloc(a)
if ((a(loc(1))) /= 1) then
    print *, "FAIL: want [1] got [", a(loc(1)), "]"
    stop 1
end if
end program t
