! vybe-test: fortran/namespace_collisions_in_scope/test_namespace_collisions_in_scope_resolves_inner_identifier
! origin: languages/fortran/tests/fortran/test_namespace_collisions_in_scope.rs

program test_namespace_collisions_in_scope
    integer :: value
    value = 1
    block
        integer :: value
        value = 3
        if ((value) /= 3) then
    print *, "FAIL: want [3] got [", value, "]"
    stop 1
end if
    end block
    if ((value) /= 1) then
    print *, "FAIL: want [1] got [", value, "]"
    stop 1
end if
end program test_namespace_collisions_in_scope
