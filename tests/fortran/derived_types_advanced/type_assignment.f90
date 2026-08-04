! vybe-test: fortran/derived_types_advanced/type_assignment
! origin: languages/fortran/tests/fortran/test_derived_types_advanced.rs

program test
    type :: Box
        integer :: width, height
    end type Box
    type(Box) :: a, b
    a%width = 10
    a%height = 20
    b = a
    print *, b%width
end program test
