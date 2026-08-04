! vybe-test: fortran/formatting/fmt_label_20
! origin: languages/fortran/tests/fortran/test_formatting.rs
program p
integer::x=1
write(*,100) x
100 format(I3)
end program p
