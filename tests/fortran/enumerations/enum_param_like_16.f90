! vybe-test: fortran/enumerations/enum_param_like_16
! origin: languages/fortran/tests/fortran/test_enumerations.rs
enum, bind(c)
enumerator :: a=1, b=2
end enum
program p
integer, parameter :: x = a
print *, x
end program p
