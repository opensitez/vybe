! vybe-test: fortran/coarray_synchronization_patterns/test_coarray_synchronization_patterns_sync_all_anchor
! origin: languages/fortran/tests/fortran/test_coarray_synchronization_patterns.rs

program test_coarray_synchronization_patterns
    integer :: value
    value = 11
    sync all
    if ((value) /= 11) then
    print *, "FAIL: want [11] got [", value, "]"
    stop 1
end if
end program test_coarray_synchronization_patterns
