! vybe-test: fortran/array_reduction_extended/any_logical_all_false
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
logical :: m(3) = [.false., .false., .false.]
if ((any(m)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", any(m), "]"
    stop 1
end if
end program t
