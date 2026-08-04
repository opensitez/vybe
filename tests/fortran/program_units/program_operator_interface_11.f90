! vybe-test: fortran/program_units/program_operator_interface_11
! origin: languages/fortran/tests/fortran/test_program_units.rs
module m
interface operator(+)
module procedure addi
end interface
contains
integer function addi(a,b)
integer :: a,b
addi = a+b
end function addi
end module m
