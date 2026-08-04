! vybe-test: fortran/pointers/nullify_basic
! origin: languages/fortran/tests/fortran/test_pointers.rs

program test
    integer, pointer :: p => null()
    integer, target :: x = 5
    p => x
    nullify(p)
    print *, associated(p)
end program test
