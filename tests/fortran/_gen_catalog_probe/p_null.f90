! vybe-test: fortran/_gen_catalog_probe/p_null
! origin: languages/fortran/tests/fortran/_gen_catalog_probe.rs
program t
use iso_c_binding
print *, merge(1,0,c_associated(c_null_ptr))
end program t
