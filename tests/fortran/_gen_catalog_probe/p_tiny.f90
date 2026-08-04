! vybe-test: fortran/_gen_catalog_probe/p_tiny
! origin: languages/fortran/tests/fortran/_gen_catalog_probe.rs
program t
print *, merge(1,0,tiny(1.0)>0.0)
end program t
