! vybe-test: fortran/hollerith/h_format_inside_repeat
! origin: languages/fortran/tests/fortran/test_hollerith.rs

program test
    write(*, 10) 1, 2, 3
10  format(3(2Hv=, I2, 1H ))
end program test
