! vybe-test: fortran/_gen_catalog_probe/p_nearest
! origin: languages/fortran/tests/fortran/_gen_catalog_probe.rs
program t
print *, merge(1,0,nearest(1.0,1.0)>1.0)
end program t
