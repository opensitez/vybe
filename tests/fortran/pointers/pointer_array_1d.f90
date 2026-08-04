! vybe-test: fortran/pointers/pointer_array_1d
! origin: languages/fortran/tests/fortran/test_pointers.rs

program test
    integer, target :: a(5) = [10, 20, 30, 40, 50]
    integer, pointer :: p(:)
    p => a
    print *, p(3)
end program test
