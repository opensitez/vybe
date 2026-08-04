! vybe-test: fortran/pointers/associated_after_target
! origin: languages/fortran/tests/fortran/test_pointers.rs

program test
    integer, target :: x = 5
    integer, pointer :: p
    p => x
    print *, associated(p)
end program test
