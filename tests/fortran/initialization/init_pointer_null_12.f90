! vybe-test: fortran/initialization/init_pointer_null_12
! origin: languages/fortran/tests/fortran/test_initialization.rs
program p
integer,pointer::p=>null()
end program p
