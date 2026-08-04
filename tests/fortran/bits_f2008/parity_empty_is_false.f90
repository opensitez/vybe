! vybe-test: fortran/bits_f2008/parity_empty_is_false
! origin: languages/fortran/tests/fortran/test_bits_f2008.rs

program test
    logical :: a(0)
    print *, parity(a)
end program test
