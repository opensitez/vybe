! vybe-test: fortran/generic_interfaces/gen_if_06
! origin: languages/fortran/tests/fortran/test_generic_interfaces.rs
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
