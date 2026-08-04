! vybe-test: fortran/coarrays/event_post_basic
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    use iso_fortran_env
    type(event_type) :: ev[*]
    if (this_image() == 1) then
        event post(ev[2])
    else if (this_image() == 2) then
        event wait(ev)
        print *, 'event received'
    end if
end program test
