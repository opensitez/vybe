! vybe-test: fortran/_gen_catalog_probe/p_ubound
! origin: languages/fortran/tests/fortran/_gen_catalog_probe.rs
program t
integer :: a(3,4)
print *, ubound(a,1)
print *, ubound(a,2)
end program t
