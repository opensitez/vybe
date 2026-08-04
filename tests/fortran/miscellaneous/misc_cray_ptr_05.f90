! vybe-test: fortran/miscellaneous/misc_cray_ptr_05
! origin: languages/fortran/tests/fortran/test_miscellaneous.rs
program p
pointer (p, x)
integer x
print *,1
end program p
