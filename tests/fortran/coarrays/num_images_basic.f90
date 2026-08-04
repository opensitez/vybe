! vybe-test: fortran/coarrays/num_images_basic
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    if ((num_images()) /= 1) then
    print *, "FAIL: want [1] got [", num_images(), "]"
    stop 1
end if
end program test
