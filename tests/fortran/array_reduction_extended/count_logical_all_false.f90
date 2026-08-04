! vybe-test: fortran/array_reduction_extended/count_logical_all_false
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
logical :: m(4) = [.false., .false., .false., .false.]
if ((count(m)) /= 0) then
    print *, "FAIL: want [0] got [", count(m), "]"
    stop 1
end if
end program t
