! vybe-test: fortran/_gen_catalog_probe/p_sumdim
! origin: languages/fortran/tests/fortran/_gen_catalog_probe.rs
program t
integer :: m(2,3)=reshape([1,2,3,4,5,6],[2,3])
print *, sum(m,dim=2)
print *, sum(m,dim=2)
end program t
