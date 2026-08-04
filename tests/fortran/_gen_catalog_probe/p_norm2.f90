! vybe-test: fortran/_gen_catalog_probe/p_norm2
! origin: languages/fortran/tests/fortran/_gen_catalog_probe.rs
program t
print *, nint(norm2([3.0,4.0,0.0]))
end program t
