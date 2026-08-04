! vybe-test: fortran/bits_f2008/dshiftr_carries_bit
! origin: languages/fortran/tests/fortran/test_bits_f2008.rs

program test
    integer :: hi = 1
    integer :: lo = 0
    print *, dshiftr(hi, lo, 1)
end program test
