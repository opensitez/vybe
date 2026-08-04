! vybe-test: fortran/fortran2003/volatile_in_module
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

module hw_reg
    implicit none
    integer, volatile :: status_reg = 0
    integer, volatile :: data_reg = 0
end module hw_reg

program test
    use hw_reg
    status_reg = 1
    print *, status_reg
end program test
