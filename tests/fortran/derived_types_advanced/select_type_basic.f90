! vybe-test: fortran/derived_types_advanced/select_type_basic
! origin: languages/fortran/tests/fortran/test_derived_types_advanced.rs

program test
    type :: Base
        integer :: id = 0
    end type Base
    type, extends(Base) :: Child
        integer :: extra = 99
    end type Child
    class(Base), allocatable :: obj
    allocate(Child :: obj)
    select type(obj)
    type is (Child)
        print *, obj%extra
    class default
        print *, "base"
    end select
end program test
