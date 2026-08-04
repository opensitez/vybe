! vybe-test: fortran/generic_resolution/generic_resolution_03
! origin: languages/fortran/tests/fortran/test_generic_resolution.rs
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
