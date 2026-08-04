! vybe-test: fortran/coarrays/sync_memory_basic
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    integer :: x[*]
    x = 0
    sync memory
    x = 1
    if ((x) /= 1) then
    print *, "FAIL: want [1] got [", x, "]"
    stop 1
end if
end program test
