! vybe-test: fortran/pointers/nullify_multiple
! origin: languages/fortran/tests/fortran/test_pointers.rs

program test
    integer, pointer :: p => null(), q => null()
    integer, target :: x = 1, y = 2
    p => x
    q => y
    nullify(p, q)
    print *, associated(p)
end program test
