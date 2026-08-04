! vybe-test: fortran/array_locators/maxloc_1d_tie_returns_first_occurrence
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(5) = [5, 3, 5, 2, 5]
integer :: loc(1)
loc = maxloc(a)
if ((loc(1)) /= 1) then
    print *, "FAIL: want [1] got [", loc(1), "]"
    stop 1
end if
end program t
