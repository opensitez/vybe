! vybe-test: fortran/fortran2008/contiguous_pointer
! origin: languages/fortran/tests/fortran/test_fortran2008.rs

program test
    integer, target :: a(10) = [(i, i=1,10)]
    integer, pointer, contiguous :: p(:)
    p => a
    print *, p(3)
end program test
