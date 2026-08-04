! vybe-test: fortran/fortran2018/event_type_declaration
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    use iso_fortran_env
    type(event_type) :: ev[*]
    print *, 'ok'
end program test
