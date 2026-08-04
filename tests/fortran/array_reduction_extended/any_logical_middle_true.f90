! vybe-test: fortran/array_reduction_extended/any_logical_middle_true
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
logical :: m(5) = [.false., .false., .true., .false., .false.]
if ((any(m)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", any(m), "]"
    stop 1
end if
end program t
