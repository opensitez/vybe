! vybe-test: fortran/_gen_catalog_probe/p_setexp
! origin: languages/fortran/tests/fortran/_gen_catalog_probe.rs
program t
print *, nint(set_exponent(1.0,3))
end program t
