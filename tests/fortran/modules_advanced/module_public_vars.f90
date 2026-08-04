! vybe-test: fortran/modules_advanced/module_public_vars
! origin: languages/fortran/tests/fortran/test_modules_advanced.rs

module config
    implicit none
    integer, public :: max_size = 100
    real, public :: tolerance = 1.0e-6
end module config

program test
    use config
    print *, max_size
end program test
