! vybe-test: fortran/derived_type_oop_extended/compile_select_type_is_exact_base_type
! origin: languages/fortran/tests/fortran/test_derived_type_oop_extended.rs

program t
    type :: Base
        integer :: id = 5
    end type Base
    type, extends(Base) :: Child
        integer :: extra = 9
    end type Child
    class(Base), allocatable :: obj
    allocate(Base :: obj)
    select type(obj)
    type is (Base)
        print *, obj%id
    class is (Child)
        print *, obj%extra
    class default
        print *, 0
    end select
end program t
