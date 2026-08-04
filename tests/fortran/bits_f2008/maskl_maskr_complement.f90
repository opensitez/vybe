! vybe-test: fortran/bits_f2008/maskl_maskr_complement
! origin: languages/fortran/tests/fortran/test_bits_f2008.rs

program test
    integer :: n = 4
    integer :: l, r
    l = maskl(n)
    r = maskr(bit_size(0) - n)
    print *, l == r
end program test
