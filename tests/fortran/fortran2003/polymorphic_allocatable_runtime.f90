! vybe-test: fortran/fortran2003/polymorphic_allocatable_runtime
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

program test
    type :: Vehicle
        integer :: wheels = 4
    end type Vehicle
    type, extends(Vehicle) :: Bike
    end type Bike
    class(Vehicle), allocatable :: v
    allocate(Bike :: v)
    if ((v%wheels) /= 4) then
    print *, "FAIL: want [4] got [", v%wheels, "]"
    stop 1
end if
end program test
