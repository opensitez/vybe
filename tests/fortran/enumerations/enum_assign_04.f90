! vybe-test: fortran/enumerations/enum_assign_04
! origin: languages/fortran/tests/fortran/test_enumerations.rs
program p
enum, bind(c)
enumerator :: red=1
end enum
integer :: x
x = red
print *, x
end program p
