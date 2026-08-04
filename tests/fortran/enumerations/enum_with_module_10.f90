! vybe-test: fortran/enumerations/enum_with_module_10
! origin: languages/fortran/tests/fortran/test_enumerations.rs
module m
enum, bind(c)
enumerator :: a=1
end enum
end module m
