! vybe-test: fortran/coarrays/num_images_requested
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    if ((num_images(requested=.true.)) /= 1) then
    print *, "FAIL: want [1] got [", num_images(requested=.true.), "]"
    stop 1
end if
end program test
