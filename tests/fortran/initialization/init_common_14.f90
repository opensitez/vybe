! vybe-test: fortran/initialization/init_common_14
! origin: languages/fortran/tests/fortran/test_initialization.rs
program p
integer::x
common /blk/ x
x=1
end program p
