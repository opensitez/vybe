! vybe-test: fortran/coarrays/event_wait_until_count
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    use iso_fortran_env
    type(event_type) :: ev[*]
    if (this_image() == 1) then
        event wait(ev, until_count=3)
        print *, 'got 3 events'
    end if
end program test
