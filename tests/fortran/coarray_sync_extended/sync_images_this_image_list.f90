! vybe-test: fortran/coarray_sync_extended/sync_images_this_image_list
! origin: languages/fortran/tests/fortran/test_coarray_sync_extended.rs

program t
    sync images ([this_image()])
    print *, 'done'
end program t
