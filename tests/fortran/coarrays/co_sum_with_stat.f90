! vybe-test: fortran/coarrays/co_sum_with_stat
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    integer :: x = 5, stat
    call co_sum(x, stat=stat)
    if ((x) /= 5) then
    print *, "FAIL: want [5] got [", x, "]"
    stop 1
end if
end program test
