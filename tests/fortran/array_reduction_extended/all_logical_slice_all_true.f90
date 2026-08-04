! vybe-test: fortran/array_reduction_extended/all_logical_slice_all_true
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
logical :: m(6) = [.true., .true., .true., .true., .false., .true.]
if ((all(m(1:4))) .neqv. .true.) then
    print *, "FAIL: want [true] got [", all(m(1:4)), "]"
    stop 1
end if
end program t
