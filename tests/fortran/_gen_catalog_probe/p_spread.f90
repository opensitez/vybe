! vybe-test: fortran/_gen_catalog_probe/p_spread
! origin: languages/fortran/tests/fortran/_gen_catalog_probe.rs
program t
integer :: a(3)=[1,2,3]
integer :: b(2,3)
b=spread(a,dim=1,n=2)
print *, b(1,2)
print *, b(2,2)
end program t
