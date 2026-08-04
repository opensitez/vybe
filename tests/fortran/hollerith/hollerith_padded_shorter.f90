! vybe-test: fortran/hollerith/hollerith_padded_shorter
! origin: languages/fortran/tests/fortran/test_hollerith.rs

program test
    integer :: w
    w = 2Hhi
    print *, 'ok'
end program test
