! vybe-test: fortran/_gen_catalog_probe/p_pack
! origin: languages/fortran/tests/fortran/_gen_catalog_probe.rs
program t
integer :: a(4)=[1,2,3,4]
logical :: m(4)=[.true.,.false.,.true.,.false.]
integer :: b(2)
b=pack(a,m)
print *, b(1)
print *, b(2)
end program t
