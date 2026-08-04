! vybe-test: fortran/specification_part/spec_save_04
! origin: languages/fortran/tests/fortran/test_specification_part.rs
program p
implicit none
integer, save :: x
x = 1
print *, x
end program p
