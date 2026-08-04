! vybe-test: fortran/bits_f2008/bitwise_compare_negative
! origin: languages/fortran/tests/fortran/test_bits_f2008.rs

program test
    integer :: a = -1, b = 1
    print *, bgt(a, b)
    print *, blt(b, a)
end program test
