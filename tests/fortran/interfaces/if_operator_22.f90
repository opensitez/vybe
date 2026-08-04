! vybe-test: fortran/interfaces/if_operator_22
! origin: languages/fortran/tests/fortran/test_interfaces.rs
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
