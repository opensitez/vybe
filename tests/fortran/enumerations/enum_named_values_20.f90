! vybe-test: fortran/enumerations/enum_named_values_20
! origin: languages/fortran/tests/fortran/test_enumerations.rs
enum, bind(c)
enumerator :: sunday=0, monday=1
end enum
