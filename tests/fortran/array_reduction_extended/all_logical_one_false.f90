! vybe-test: fortran/array_reduction_extended/all_logical_one_false
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
logical :: m(4) = [.true., .true., .false., .true.]
if ((all(m)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", all(m), "]"
    stop 1
end if
end program t
