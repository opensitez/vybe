! vybe-test: fortran/bits_f2008/parity_with_dim
! origin: languages/fortran/tests/fortran/test_bits_f2008.rs

program test
    logical :: m(2,3) = reshape([.true.,.false.,.true.,.false.,.true.,.false.],[2,3])
    logical :: row_parity(3)
    row_parity = parity(m, dim=1)
    print *, row_parity(1)
end program test
