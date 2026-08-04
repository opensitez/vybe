! vybe-test: fortran/enumerations/enum_assign_04
! origin: languages/fortran/tests/fortran/test_enumerations.rs
enum, bind(c)
enumerator :: red=1
end enum
program p
integer :: x
x = red
print *, x
end program p
