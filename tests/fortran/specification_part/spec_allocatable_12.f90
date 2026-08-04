! vybe-test: fortran/specification_part/spec_allocatable_12
! origin: languages/fortran/tests/fortran/test_specification_part.rs
program p
implicit none
integer, allocatable :: a(:)
print *, 1
end program p
