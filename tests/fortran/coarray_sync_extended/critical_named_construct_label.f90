! vybe-test: fortran/coarray_sync_extended/critical_named_construct_label
! origin: languages/fortran/tests/fortran/test_coarray_sync_extended.rs

program t
    integer :: slot[*]
    slot = 0
    sync all
    guard: critical
        slot[1] = slot[1] + this_image()
    end critical guard
    sync all
    if (this_image() == 1) print *, slot
end program t
