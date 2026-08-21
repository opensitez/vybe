! vybe-test: fortran/allocation_status/allocated_false
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

program test
    integer, allocatable :: x(:)
    print *, allocated(x)
end program test
