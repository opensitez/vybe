! vybe-test: fortran/hollerith/hollerith_assigned_to_real
! origin: languages/fortran/tests/fortran/test_hollerith.rs

program test
    real :: tag
    tag = 4HTEST
    print *, 'ok'
end program test
