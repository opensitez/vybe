! vybe-test: fortran/array_locators/maxloc_1d_negative_values_peak_at_start
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(5) = [-1, -5, -2, -5, -3]
integer :: loc(1)
loc = maxloc(a)
if ((loc(1)) /= 1) then
    print *, "FAIL: want [1] got [", loc(1), "]"
    stop 1
end if
end program t
