! vybe-test: fortran/enumerations/enum_print_15
! origin: languages/fortran/tests/fortran/test_enumerations.rs
enum, bind(c)
enumerator :: a=1
end enum
program p
print *, a
end program p
