! vybe-test: fortran/array_locators/findloc_back_with_mask_restricted_tail
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(6) = [1, 2, 1, 2, 1, 2]
logical :: mask(6) = [.true., .true., .true., .true., .true., .false.]
integer :: loc(1)
loc = findloc(a, 1, back=.true., mask=mask)
if ((loc(1)) /= 5) then
    print *, "FAIL: want [5] got [", loc(1), "]"
    stop 1
end if
end program t
