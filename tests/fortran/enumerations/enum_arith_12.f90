! vybe-test: fortran/enumerations/enum_arith_12
! origin: languages/fortran/tests/fortran/test_enumerations.rs
program p
enum, bind(c)
enumerator :: a=1, b=2
end enum
print *, a+b
end program p
