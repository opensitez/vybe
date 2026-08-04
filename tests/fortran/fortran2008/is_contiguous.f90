! vybe-test: fortran/fortran2008/is_contiguous
! origin: languages/fortran/tests/fortran/test_fortran2008.rs

program test
    integer, target :: a(10)
    integer, pointer :: p(:)
    p => a
    print *, is_contiguous(p)
end program test
