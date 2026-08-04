! vybe-test: fortran/array_locators/maxloc_1d_unimodal_peak_at_four
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(7) = [1, 3, 5, 7, 6, 4, 2]
integer :: loc(1)
loc = maxloc(a)
if ((loc(1)) /= 4) then
    print *, "FAIL: want [4] got [", loc(1), "]"
    stop 1
end if
end program t
