! vybe-test: fortran/formatting/fmt_internal_16
! origin: languages/fortran/tests/fortran/test_formatting.rs
program p
character(len=20)::buf
write(buf,'(I3)') 42
print *, trim(buf)
end program p
