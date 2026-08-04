! vybe-test: fortran/formatting/fmt_read_17
! origin: languages/fortran/tests/fortran/test_formatting.rs
program p
character(len=20)::buf='42'
integer::x
read(buf,'(I2)') x
print *,x
end program p
