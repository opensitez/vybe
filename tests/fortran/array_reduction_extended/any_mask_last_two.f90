! vybe-test: fortran/array_reduction_extended/any_mask_last_two
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
logical :: m(4) = [.false., .false., .true., .true.]
if ((any(m)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", any(m), "]"
    stop 1
end if
end program t
