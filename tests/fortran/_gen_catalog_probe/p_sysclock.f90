! vybe-test: fortran/_gen_catalog_probe/p_sysclock
! origin: languages/fortran/tests/fortran/_gen_catalog_probe.rs
program t
integer :: c,r,m
call system_clock(c,r,m)
print *, merge(1,0,m>0)
end program t
