! vybe-test: fortran/hollerith/hollerith_multiword
! origin: languages/fortran/tests/fortran/test_hollerith.rs

program test
    write(*, 100)
100 format(13Hhello, world!)
end program test
