! vybe-test: fortran/enumerations/enum_expr_17
! origin: languages/fortran/tests/fortran/test_enumerations.rs
enum, bind(c)
enumerator :: a=1, b=2
end enum
program p
integer :: x
x = a * b
print *, x
end program p
