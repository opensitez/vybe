! vybe-test: fortran/coarray_sync_extended/sync_images_int_array_target
! origin: languages/fortran/tests/fortran/test_coarray_sync_extended.rs
program t
integer :: peers(1)
peers(1) = this_image()
sync images (peers)
if (trim('peers ok') /= "peers ok") then
    print *, "FAIL: want [peers ok] got [", 'peers ok', "]"
    stop 1
end if
end program t
