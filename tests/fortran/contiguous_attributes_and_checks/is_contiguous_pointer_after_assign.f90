! vybe-test: fortran/contiguous_attributes_and_checks/is_contiguous_pointer_after_assign
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    real, target :: a(10)
    real, pointer :: p(:)
    p => a
    print *, is_contiguous(p)
end program test
