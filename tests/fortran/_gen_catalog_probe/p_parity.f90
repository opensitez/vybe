! vybe-test: fortran/_gen_catalog_probe/p_parity
! origin: languages/fortran/tests/fortran/_gen_catalog_probe.rs
program t
logical :: a(2)=[.true.,.true.]
print *, merge(1,0,parity(a))
end program t
