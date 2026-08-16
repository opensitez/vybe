! vybe-test: fortran/enumerations/enum_kind_int_19
! origin: languages/fortran/tests/fortran/test_enumerations.rs
program p
enum, bind(c)
enumerator :: a=1
end enum
integer :: x
x = a
end program p
