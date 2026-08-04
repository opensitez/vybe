! vybe-test: fortran/pointer_component_disambiguation/test_pointer_component_disambiguation_resolves_target_member
! origin: languages/fortran/tests/fortran/test_pointer_component_disambiguation.rs

program test_pointer_component_disambiguation
    type :: node
        integer :: a
    end type

    type(node), target :: n
    type(node), pointer :: p
    n%a = 8
    p => n
    if ((p%a) /= 8) then
    print *, "FAIL: want [8] got [", p%a, "]"
    stop 1
end if
end program test_pointer_component_disambiguation
