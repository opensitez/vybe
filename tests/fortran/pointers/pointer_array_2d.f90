! vybe-test: fortran/pointers/pointer_array_2d
! origin: languages/fortran/tests/fortran/test_pointers.rs

program test
    integer, target :: m(3,3)
    integer, pointer :: p(:,:)
    m = 0
    m(2,2) = 42
    p => m
    print *, p(2,2)
end program test
