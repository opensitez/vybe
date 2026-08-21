! vybe-test: fortran/iso_fortran_env/iso_fortran_env_units
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

program test
    use iso_fortran_env
    write(output_unit, *) 'stdout'
    write(error_unit, *) 'stderr'
end program test
