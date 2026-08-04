! vybe-test: fortran/array_locators/maxloc_1d_sparse_peak_at_four
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(5) = [0, 0, 0, 5, 0]
integer :: loc(1)
loc = maxloc(a)
if ((loc(1)) /= 4) then
    print *, "FAIL: want [4] got [", loc(1), "]"
    stop 1
end if
end program t
