! vybe-test: fortran/pointers/pointer_array_slice
! origin: languages/fortran/tests/fortran/test_pointers.rs

program test
    integer, target :: a(6) = [1, 2, 3, 4, 5, 6]
    integer, pointer :: p(:)
    p => a(2:5)
    print *, p(1)
end program test
