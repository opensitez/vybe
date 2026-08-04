! vybe-test: fortran/pointers/pointer_reassociate
! origin: languages/fortran/tests/fortran/test_pointers.rs

program test
    integer, target :: a = 1, b = 2
    integer, pointer :: p
    p => a
    p => b
    print *, p
end program test
