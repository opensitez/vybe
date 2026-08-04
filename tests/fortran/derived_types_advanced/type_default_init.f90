! vybe-test: fortran/derived_types_advanced/type_default_init
! origin: languages/fortran/tests/fortran/test_derived_types_advanced.rs

program test
    type :: Config
        integer :: timeout = 30
        real :: threshold = 0.01
        logical :: debug = .false.
    end type Config
    type(Config) :: cfg
    print *, cfg%timeout
end program test
