! vybe-test: fortran/hollerith/hollerith_single_char
! origin: languages/fortran/tests/fortran/test_hollerith.rs

program test
    write(*, 100)
100 format(1Hx)
end program test
