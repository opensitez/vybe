! vybe-test: fortran/hollerith/hollerith_space
! origin: languages/fortran/tests/fortran/test_hollerith.rs

program test
    integer :: a = 1, b = 2
    write(*, 100) a, b
100 format(I3, 1H , I3)
end program test
