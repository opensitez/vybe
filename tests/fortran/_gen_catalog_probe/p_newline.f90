! vybe-test: fortran/_gen_catalog_probe/p_newline
! origin: languages/fortran/tests/fortran/_gen_catalog_probe.rs
program t
character(len=1)::nl
nl=new_line('a')
print *, ichar(nl)
end program t
