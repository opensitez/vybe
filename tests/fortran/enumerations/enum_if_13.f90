! vybe-test: fortran/enumerations/enum_if_13
! origin: languages/fortran/tests/fortran/test_enumerations.rs
enum, bind(c)
enumerator :: a=1
end enum
program p
if (a == 1) print *,1
end program p
