! vybe-test: fortran/miscellaneous/misc_exec_seq_01
! origin: languages/fortran/tests/fortran/test_miscellaneous.rs
program p
integer::x
x=1
x=x+1
print *,x
end program p
