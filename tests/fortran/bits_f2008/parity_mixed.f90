! vybe-test: fortran/bits_f2008/parity_mixed
! origin: languages/fortran/tests/fortran/test_bits_f2008.rs

program test
    logical :: a(4) = [.true., .false., .true., .false.]
    print *, parity(a)
end program test
