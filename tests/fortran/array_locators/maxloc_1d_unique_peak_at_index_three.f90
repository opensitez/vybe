! vybe-test: fortran/array_locators/maxloc_1d_unique_peak_at_index_three
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(5) = [3, 1, 9, 1, 5]
integer :: loc(1)
loc = maxloc(a)
if ((loc(1)) /= 3) then
    print *, "FAIL: want [3] got [", loc(1), "]"
    stop 1
end if
end program t
