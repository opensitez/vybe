! vybe-test: fortran/array_reduction_extended/any_logical_slice_middle
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
logical :: m(5) = [.true., .false., .true., .false., .true.]
if ((any(m(2:4))) .neqv. .true.) then
    print *, "FAIL: want [true] got [", any(m(2:4)), "]"
    stop 1
end if
end program t
