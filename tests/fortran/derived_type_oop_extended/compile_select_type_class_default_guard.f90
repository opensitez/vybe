! vybe-test: fortran/derived_type_oop_extended/compile_select_type_class_default_guard
! origin: languages/fortran/tests/fortran/test_derived_type_oop_extended.rs

program t
    type :: Root
        integer :: tag = 1
    end type Root
    type, extends(Root) :: Leaf
        integer :: payload = 4
    end type Leaf
    class(Root), allocatable :: node
    allocate(Root :: node)
    select type(node)
    class is (Leaf)
        print *, node%payload
    type is (Root)
        print *, node%tag
    class default
        print *, 0
    end select
end program t
