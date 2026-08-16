! vybe-test: fortran/enumerations/enum_print_15
! origin: languages/fortran/tests/fortran/test_enumerations.rs
program p
enum, bind(c)
enumerator :: a=1
end enum
print *, a
end program p
