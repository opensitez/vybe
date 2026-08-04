! vybe-test: fortran/coarrays/sync_images_specific
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    if (this_image() == 1) then
        sync images ([2, 3])
    else
        sync images ([1])
    end if
    print *, 'synced'
end program test
