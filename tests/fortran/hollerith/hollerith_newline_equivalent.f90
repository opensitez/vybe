! vybe-test: fortran/hollerith/hollerith_newline_equivalent
! origin: languages/fortran/tests/fortran/test_hollerith.rs

program test
    write(*, 100)
100 format(4Hline)
    write(*, 200)
200 format(4Htwo!)
end program test
