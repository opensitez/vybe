! vybe-test: fortran/pointers/pointer_field
! origin: languages/fortran/tests/fortran/test_pointers.rs

program test
    type :: Node
        integer :: value
        type(Node), pointer :: next => null()
    end type Node
    type(Node) :: n
    n%value = 1
    print *, n%value
end program test
