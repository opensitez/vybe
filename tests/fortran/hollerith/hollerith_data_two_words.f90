! vybe-test: fortran/hollerith/hollerith_data_two_words
! origin: languages/fortran/tests/fortran/test_hollerith.rs

program test
    integer :: w1, w2
    data w1 /4HTEST/, w2 /4HDATA/
    print *, 'ok'
end program test
