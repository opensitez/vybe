! vybe-test: fortran/specification_part/spec_intrinsic_09
! origin: languages/fortran/tests/fortran/test_specification_part.rs
program p
implicit none
intrinsic abs
print *, abs(-1)
end program p
