! vybe-test: fortran/_gen_catalog_probe/p_merge_arr
! origin: languages/fortran/tests/fortran/_gen_catalog_probe.rs
program t
integer :: a(3)=[1,2,3]
integer :: b(3)=[9,8,7]
logical :: m(3)=[.true.,.false.,.true.]
print *, merge(a,b,m)
print *, merge(a,b,m)
end program t
