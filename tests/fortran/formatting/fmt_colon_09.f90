! vybe-test: fortran/formatting/fmt_colon_09
! origin: languages/fortran/tests/fortran/test_formatting.rs
program p
integer::x=1
write(*,'(I2,:,I2)') x
end program p
