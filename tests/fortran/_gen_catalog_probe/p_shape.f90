! vybe-test: fortran/_gen_catalog_probe/p_shape
! origin: languages/fortran/tests/fortran/_gen_catalog_probe.rs
program t
integer :: a(2,3)
print *, shape(a,1)
print *, shape(a,2)
end program t
