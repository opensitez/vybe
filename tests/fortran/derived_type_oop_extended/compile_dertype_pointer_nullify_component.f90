! vybe-test: fortran/derived_type_oop_extended/compile_dertype_pointer_nullify_component
! origin: languages/fortran/tests/fortran/test_derived_type_oop_extended.rs

program t
    type :: Link
        integer :: value = 0
        type(Link), pointer :: next => null()
    end type Link
    type(Link), target :: head, tail
    head%value = 1
    tail%value = 2
    head%next => tail
    nullify(head%next)
    print *, head%value
end program t
