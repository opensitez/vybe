! vybe-test: fortran/enumerations/enum_arith_12
! origin: languages/fortran/tests/fortran/test_enumerations.rs
enum, bind(c)
enumerator :: a=1, b=2
end enum
program p
print *, a+b
end program p
