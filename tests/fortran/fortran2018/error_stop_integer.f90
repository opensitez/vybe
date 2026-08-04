! vybe-test: fortran/fortran2018/error_stop_integer
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    logical :: ok = .true.
    if (.not. ok) error stop 1
    print *, 'fine'
end program test
