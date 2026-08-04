! vybe-test: fortran/hollerith/hollerith_in_format
! origin: languages/fortran/tests/fortran/test_hollerith.rs

program test
    write(*, 100)
100 format(5Hhello)
end program test
