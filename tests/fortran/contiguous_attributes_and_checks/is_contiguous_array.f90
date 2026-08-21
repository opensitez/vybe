! vybe-test: fortran/contiguous_attributes_and_checks/is_contiguous_array
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    real :: a(10)
    print *, is_contiguous(a)
end program test
