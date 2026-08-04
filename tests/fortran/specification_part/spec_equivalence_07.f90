! vybe-test: fortran/specification_part/spec_equivalence_07
! origin: languages/fortran/tests/fortran/test_specification_part.rs
program p
implicit none
integer :: a, b
equivalence (a, b)
print *, 1
end program p
