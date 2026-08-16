! vybe-test: fortran/bits_f2008/trailz_and_leadz_together
! origin: languages/fortran/tests/fortran/test_bits_f2008.rs

program test
    integer :: x = 16
    print *, leadz(x)
    print *, trailz(x)
    print *, popcnt(x)
end program test
