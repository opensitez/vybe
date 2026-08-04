! vybe-test: fortran/coarray_sync_extended/sync_images_bracket_list_literal
! origin: languages/fortran/tests/fortran/test_coarray_sync_extended.rs
program t
sync images ([1])
if (trim('self sync') /= "self sync") then
    print *, "FAIL: want [self sync] got [", 'self sync', "]"
    stop 1
end if
end program t
