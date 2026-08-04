! vybe-test: fortran/_gen_catalog_probe/p_dim
! origin: languages/fortran/tests/fortran/_gen_catalog_probe.rs
program t
print *, nint(dim(10.5,3.2))
end program t
