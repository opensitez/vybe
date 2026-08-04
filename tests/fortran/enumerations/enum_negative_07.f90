! vybe-test: fortran/enumerations/enum_negative_07
! origin: languages/fortran/tests/fortran/test_enumerations.rs
enum, bind(c)
enumerator :: a=-1, b=0
end enum
