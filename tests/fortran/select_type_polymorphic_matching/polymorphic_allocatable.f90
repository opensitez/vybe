! vybe-test: fortran/select_type_polymorphic_matching/polymorphic_allocatable
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

program test
    type :: Vehicle
        integer :: wheels = 4
    end type Vehicle
    type, extends(Vehicle) :: Bike
    end type Bike
    class(Vehicle), allocatable :: v
    allocate(Bike :: v)
    print *, v%wheels
end program test
