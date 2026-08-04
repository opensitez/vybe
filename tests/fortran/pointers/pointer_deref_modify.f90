! vybe-test: fortran/pointers/pointer_deref_modify
! origin: languages/fortran/tests/fortran/test_pointers.rs

program test
    integer, target :: x = 10
    integer, pointer :: p
    p => x
    p = 99
    print *, x
end program test
