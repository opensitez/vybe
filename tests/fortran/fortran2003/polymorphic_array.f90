! vybe-test: fortran/fortran2003/polymorphic_array
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

program test
    type :: Base
        integer :: id = 0
    end type Base
    class(Base), allocatable :: arr(:)
    allocate(Base :: arr(3))
    arr(1)%id = 1
    print *, arr(1)%id
end program test
