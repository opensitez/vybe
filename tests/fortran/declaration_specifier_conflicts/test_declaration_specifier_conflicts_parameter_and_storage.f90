! vybe-test: fortran/declaration_specifier_conflicts/test_declaration_specifier_conflicts_parameter_and_storage
! origin: languages/fortran/tests/fortran/test_declaration_specifier_conflicts.rs

program test_declaration_specifier_conflicts
    integer, parameter :: a = 7
    integer, target :: b
    integer, pointer :: p
    b = a
    p => b
    if ((a) /= 7) then
    print *, "FAIL: want [7] got [", a, "]"
    stop 1
end if
    if ((p) /= 7) then
    print *, "FAIL: want [7] got [", p, "]"
    stop 1
end if
end program test_declaration_specifier_conflicts
