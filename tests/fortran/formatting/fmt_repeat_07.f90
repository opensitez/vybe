! vybe-test: fortran/formatting/fmt_repeat_07
! origin: languages/fortran/tests/fortran/test_formatting.rs
program p
integer::x=1
write(*,'(3I2)') x,x,x
end program p
