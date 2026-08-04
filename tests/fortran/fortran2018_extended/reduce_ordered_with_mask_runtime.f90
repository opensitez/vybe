! vybe-test: fortran/fortran2018_extended/reduce_ordered_with_mask_runtime
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs

program t
    integer :: a(6) = [6, 5, 4, 3, 2, 1]
    logical :: mask(6) = [.true., .false., .true., .false., .true., .false.]
    integer :: r
    r = reduce(a, operator(+), mask=mask, ordered=.true.)
    if ((r) /= 12) then
    print *, "FAIL: want [12] got [", r, "]"
    stop 1
end if
end program t
