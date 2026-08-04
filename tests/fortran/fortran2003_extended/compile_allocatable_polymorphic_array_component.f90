! vybe-test: fortran/fortran2003_extended/compile_allocatable_polymorphic_array_component
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs

program t
    type :: Part
        integer :: id = 0
    end type Part
    type :: Assembly
        class(Part), allocatable :: parts(:)
    end type Assembly
    type(Assembly) :: a
    allocate(Part :: a%parts(2))
    a%parts(1)%id = 3
  a%parts(2)%id = 4
    print *, a%parts(1)%id + a%parts(2)%id
end program t
