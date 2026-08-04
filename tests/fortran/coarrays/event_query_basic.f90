! vybe-test: fortran/coarrays/event_query_basic
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    use iso_fortran_env
    type(event_type) :: ev[*]
    integer :: count
    event query(ev, count)
    print *, count
end program test
