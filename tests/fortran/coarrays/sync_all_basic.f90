! vybe-test: fortran/coarrays/sync_all_basic
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    integer :: x[*]
    x = this_image() * 10
    sync all
    if ((x) /= 10) then
    print *, "FAIL: want [10] got [", x, "]"
    stop 1
end if
end program test
