! vybe-test: fortran/fortran2018_extended/error_stop_character_message
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs

program t
    logical :: ok = .true.
    if (.not. ok) error stop 'aborted'
    print *, 'fine'
end program t
