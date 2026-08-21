! vybe-test: fortran/contiguous_attributes_and_checks/is_contiguous_non_unit_stride
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    real, target :: a(10)
    real, pointer :: p(:)
    p => a(1:10:2)
    print *, is_contiguous(p)
end program test
