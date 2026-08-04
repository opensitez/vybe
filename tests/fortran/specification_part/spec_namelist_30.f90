! vybe-test: fortran/specification_part/spec_namelist_30
! origin: languages/fortran/tests/fortran/test_specification_part.rs
program p
implicit none
integer :: x
namelist /n1/ x
print *, 1
end program p
