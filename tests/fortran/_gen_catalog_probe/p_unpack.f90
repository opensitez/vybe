! vybe-test: fortran/_gen_catalog_probe/p_unpack
! origin: languages/fortran/tests/fortran/_gen_catalog_probe.rs
program t
integer :: a(2)=[7,9]
logical :: m(4)=[.true.,.false.,.true.,.false.]
integer :: b(4)
b=unpack(a,m,0)
print *, b(1)
print *, b(3)
end program t
