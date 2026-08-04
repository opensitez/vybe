! vybe-test: fortran/hollerith/hollerith_assigned_to_integer
! origin: languages/fortran/tests/fortran/test_hollerith.rs

program test
    integer :: tag
    tag = 4HTEST
    print *, 'ok'
end program test
