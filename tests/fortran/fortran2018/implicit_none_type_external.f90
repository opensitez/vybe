! vybe-test: fortran/fortran2018/implicit_none_type_external
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    implicit none (type, external)
    integer :: x = 42
    print *, x
end program test
