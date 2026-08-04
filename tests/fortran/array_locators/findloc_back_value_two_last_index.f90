! vybe-test: fortran/array_locators/findloc_back_value_two_last_index
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(6) = [1, 2, 1, 2, 1, 2]
integer :: loc(1)
loc = findloc(a, 2, back=.true.)
if ((loc(1)) /= 6) then
    print *, "FAIL: want [6] got [", loc(1), "]"
    stop 1
end if
end program t
