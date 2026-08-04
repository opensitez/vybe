! vybe-test: fortran/hollerith/hollerith_in_data
! origin: languages/fortran/tests/fortran/test_hollerith.rs

program test
    integer :: word
    data word /4HABCD/
    print *, 'ok'
end program test
