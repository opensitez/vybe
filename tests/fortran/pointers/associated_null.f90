! vybe-test: fortran/pointers/associated_null
! origin: languages/fortran/tests/fortran/test_pointers.rs

program test
    integer, pointer :: p => null()
    print *, associated(p)
end program test
