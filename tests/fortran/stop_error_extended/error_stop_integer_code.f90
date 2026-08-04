! vybe-test: fortran/stop_error_extended/error_stop_integer_code
! origin: languages/fortran/tests/fortran/test_stop_error_extended.rs

program t
    logical :: ok = .true.
    if (.not. ok) error stop 1
    print *, 'fine'
end program t
