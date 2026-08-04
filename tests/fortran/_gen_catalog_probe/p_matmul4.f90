! vybe-test: fortran/_gen_catalog_probe/p_matmul4
! origin: languages/fortran/tests/fortran/_gen_catalog_probe.rs
program t
integer :: a(4,4),b(4,4),c(4,4)
a=reshape([(i,i=1,16)],[4,4])
b=0;b(1,1)=1;b(2,2)=1;b(3,3)=1;b(4,4)=1
c=matmul(a,b)
print *, c(1,1)
print *, c(4,4)
end program t
