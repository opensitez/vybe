! vybe-test: fortran/array_locators/maxloc_1d_trailing_nine_at_five
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(5) = [8, 2, 6, 2, 9]
integer :: loc(1)
loc = maxloc(a)
if ((loc(1)) /= 5) then
    print *, "FAIL: want [5] got [", loc(1), "]"
    stop 1
end if
end program t
