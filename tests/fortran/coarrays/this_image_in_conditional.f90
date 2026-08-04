! vybe-test: fortran/coarrays/this_image_in_conditional
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    if (this_image() == 1) then
        if (trim('image 1') /= "image 1") then
    print *, "FAIL: want [image 1] got [", 'image 1', "]"
    stop 1
end if
    end if
end program test
