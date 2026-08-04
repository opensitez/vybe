! vybe-test: fortran/miscellaneous/misc_optbarrier_15
! origin: languages/fortran/tests/fortran/test_miscellaneous.rs
program p
integer,volatile::x
x=1
print *,x
end program p
