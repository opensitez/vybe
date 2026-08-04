! vybe-test: fortran/hollerith/h_format_label_and_value
! origin: languages/fortran/tests/fortran/test_hollerith.rs

program test
    real :: x = 3.14
    write(*, 10) x
10  format(7Hresult=, F6.2)
end program test
