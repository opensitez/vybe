! vybe-test: fortran/hollerith/hollerith_with_integer
! origin: languages/fortran/tests/fortran/test_hollerith.rs

program test
    integer :: n = 42
    write(*, 100) n
100 format(5Hval= , I4)
end program test
