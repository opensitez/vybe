! vybe-test: fortran/array_locators/minloc_1d_leading_two_at_index_two
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(5) = [8, 2, 6, 2, 9]
integer :: loc(1)
loc = minloc(a)
if ((loc(1)) /= 2) then
    print *, "FAIL: want [2] got [", loc(1), "]"
    stop 1
end if
end program t
