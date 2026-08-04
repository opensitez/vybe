! vybe-test: fortran/coarrays/sync_images_star
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    sync images (*)
    if (trim('done') /= "done") then
    print *, "FAIL: want [done] got [", 'done', "]"
    stop 1
end if
end program test
