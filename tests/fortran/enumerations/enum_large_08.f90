! vybe-test: fortran/enumerations/enum_large_08
! origin: languages/fortran/tests/fortran/test_enumerations.rs
enum, bind(c)
enumerator :: big=1000
end enum
