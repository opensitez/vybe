! vybe-test: fortran/iso_fortran_env/iso_fortran_env_compiler_version
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

program test
    use iso_fortran_env
    print *, compiler_version()
    print *, compiler_options()
end program test
