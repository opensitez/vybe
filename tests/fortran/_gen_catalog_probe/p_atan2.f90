! vybe-test: fortran/_gen_catalog_probe/p_atan2
! origin: languages/fortran/tests/fortran/_gen_catalog_probe.rs
program t
print *, nint(atan2(2.0,2.0)*1000)
end program t
