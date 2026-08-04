! vybe-test: fortran/array_reduction_extended/count_logical_three_of_five
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
logical :: m(5) = [.true., .false., .true., .false., .true.]
if ((count(m)) /= 3) then
    print *, "FAIL: want [3] got [", count(m), "]"
    stop 1
end if
end program t
