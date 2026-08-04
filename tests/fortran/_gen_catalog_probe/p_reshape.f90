! vybe-test: fortran/_gen_catalog_probe/p_reshape
! origin: languages/fortran/tests/fortran/_gen_catalog_probe.rs
program t
integer :: a(6)=[1,2,3,4,5,6]
integer :: b(2,3)
b=reshape(a,[2,3],order=[2,1])
print *, b(1,1)
print *, b(2,1)
end program t
