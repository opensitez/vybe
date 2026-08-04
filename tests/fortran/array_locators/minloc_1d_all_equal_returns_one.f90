! vybe-test: fortran/array_locators/minloc_1d_all_equal_returns_one
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(3) = [4, 4, 4]
integer :: loc(1)
loc = minloc(a)
if ((loc(1)) /= 1) then
    print *, "FAIL: want [1] got [", loc(1), "]"
    stop 1
end if
end program t
