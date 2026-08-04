! vybe-test: fortran/pointers/associated_with_target
! origin: languages/fortran/tests/fortran/test_pointers.rs

program test
    integer, target :: x = 1, y = 2
    integer, pointer :: p
    p => x
    print *, associated(p, x)
    print *, associated(p, y)
end program test
