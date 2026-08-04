! vybe-test: fortran/array_locators/findloc_forward_first_occurrence
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(6) = [1, 2, 1, 2, 1, 2]
integer :: loc(1)
loc = findloc(a, 1)
if ((loc(1)) /= 1) then
    print *, "FAIL: want [1] got [", loc(1), "]"
    stop 1
end if
end program t
