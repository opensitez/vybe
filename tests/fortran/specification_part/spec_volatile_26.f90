! vybe-test: fortran/specification_part/spec_volatile_26
! origin: languages/fortran/tests/fortran/test_specification_part.rs
program p
implicit none
integer, volatile :: x
print *, 1
end program p
