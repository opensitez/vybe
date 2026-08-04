! vybe-test: fortran/generic_ambiguity/generic_ambiguity_04
! origin: languages/fortran/tests/fortran/test_generic_ambiguity.rs
module m
interface operator(+)
module procedure addi,addr
end interface
contains
integer function addi(a,b)
integer::a,b
addi=a+b
end
real function addr(a,b)
real::a,b
addr=a+b
end
end module m
