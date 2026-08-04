! vybe-test: fortran/fortran2008/ieee_exceptions
! origin: languages/fortran/tests/fortran/test_fortran2008.rs

program test
    use ieee_exceptions
    type(ieee_flag_type) :: flag
    logical :: halting
    call ieee_get_halting_mode(ieee_divide_by_zero, halting)
    print *, halting
end program test
