! vybe-test: fortran/bits_f2008/dshiftl_carries_bit
! origin: languages/fortran/tests/fortran/test_bits_f2008.rs

program test
    integer :: hi = int(z'80000000')
    integer :: lo = 0
    print *, dshiftl(hi, lo, 1)
end program test
