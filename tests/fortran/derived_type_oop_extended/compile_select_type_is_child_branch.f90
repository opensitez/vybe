! vybe-test: fortran/derived_type_oop_extended/compile_select_type_is_child_branch
! origin: languages/fortran/tests/fortran/test_derived_type_oop_extended.rs

program t
    type :: Base
        integer :: id = 0
    end type Base
    type, extends(Base) :: Child
        integer :: extra = 42
    end type Child
    class(Base), allocatable :: obj
    allocate(Child :: obj)
    select type(obj)
    type is (Child)
        print *, obj%extra
    class default
        print *, obj%id
    end select
end program t
