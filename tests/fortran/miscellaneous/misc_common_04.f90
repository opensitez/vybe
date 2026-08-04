! vybe-test: fortran/miscellaneous/misc_common_04
! origin: languages/fortran/tests/fortran/test_miscellaneous.rs
program p
integer::x
common /blk/ x
print *,x
end program p
