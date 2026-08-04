! vybe-test: fortran/specification_part/spec_common_06
! origin: languages/fortran/tests/fortran/test_specification_part.rs
program p
implicit none
integer :: x
common /blk/ x
print *, x
end program p
