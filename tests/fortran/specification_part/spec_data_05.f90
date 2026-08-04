! vybe-test: fortran/specification_part/spec_data_05
! origin: languages/fortran/tests/fortran/test_specification_part.rs
program p
implicit none
integer :: x
data x /1/
print *, x
end program p
