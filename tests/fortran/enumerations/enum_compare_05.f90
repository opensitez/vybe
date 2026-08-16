! vybe-test: fortran/enumerations/enum_compare_05
! origin: languages/fortran/tests/fortran/test_enumerations.rs
program p
enum, bind(c)
enumerator :: red=1, blue=2
end enum
logical :: l
l = red < blue
print *, l
end program p
