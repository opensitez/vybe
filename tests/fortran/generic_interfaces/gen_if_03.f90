! vybe-test: fortran/generic_interfaces/gen_if_03
! origin: languages/fortran/tests/fortran/test_generic_interfaces.rs
module m
interface operator(+)
module procedure addi
end interface
contains
integer function addi(a,b)
integer::a,b
addi=a+b
end
end module m
