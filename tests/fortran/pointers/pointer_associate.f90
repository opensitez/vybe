! vybe-test: fortran/pointers/pointer_associate
! origin: languages/fortran/tests/fortran/test_pointers.rs

program test
    integer, target :: x = 10
    integer, pointer :: p
    p => x
    print *, p
end program test
