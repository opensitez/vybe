! vybe-test: fortran/interfaces/if_operator_minus_34
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
interface operator(-)
module procedure subi
end interface
contains
integer function subi(a,b)
integer::a,b
subi=a-b
end
end module m
