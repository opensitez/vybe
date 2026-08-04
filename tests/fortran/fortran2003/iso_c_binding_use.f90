! vybe-test: fortran/fortran2003/iso_c_binding_use
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

program test
    use iso_c_binding
    integer(c_int) :: n = 42_c_int
    real(c_double) :: x = 3.14_c_double
    print *, n
    print *, x
end program test
