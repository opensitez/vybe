! vybe-test: fortran/_gen_catalog_probe/p_present
! origin: languages/fortran/tests/fortran/_gen_catalog_probe.rs
program t
call sub(1)
contains
subroutine sub(x,optional y)
integer,intent(in)::x
integer,optional,intent(in)::y
print *, merge(1,0,.not.present(y))
end subroutine sub
end program t
