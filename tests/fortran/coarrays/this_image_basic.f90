! vybe-test: fortran/coarrays/this_image_basic
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    integer :: me
    me = this_image()
    if ((me) /= 1) then
    print *, "FAIL: want [1] got [", me, "]"
    stop 1
end if
end program test
