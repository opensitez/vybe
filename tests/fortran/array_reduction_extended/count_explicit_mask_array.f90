! vybe-test: fortran/array_reduction_extended/count_explicit_mask_array
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
logical :: m(6) = [.true., .false., .true., .true., .false., .false.]
if ((count(m)) /= 3) then
    print *, "FAIL: want [3] got [", count(m), "]"
    stop 1
end if
end program t
