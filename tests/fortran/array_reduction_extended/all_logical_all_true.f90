! vybe-test: fortran/array_reduction_extended/all_logical_all_true
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
logical :: m(4) = [.true., .true., .true., .true.]
if ((all(m)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", all(m), "]"
    stop 1
end if
end program t
