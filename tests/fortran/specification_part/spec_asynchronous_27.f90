! vybe-test: fortran/specification_part/spec_asynchronous_27
! origin: languages/fortran/tests/fortran/test_specification_part.rs
program p
implicit none
integer, asynchronous :: x
print *, 1
end program p
