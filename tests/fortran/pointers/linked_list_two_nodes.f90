! vybe-test: fortran/pointers/linked_list_two_nodes
! origin: languages/fortran/tests/fortran/test_pointers.rs

program test
    type :: Node
        integer :: value
        type(Node), pointer :: next => null()
    end type Node
    type(Node), target :: n1, n2
    n1%value = 1
    n2%value = 2
    n1%next => n2
    print *, n1%next%value
end program test
