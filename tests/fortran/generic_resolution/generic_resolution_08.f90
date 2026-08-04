! vybe-test: fortran/generic_resolution/generic_resolution_08
! origin: languages/fortran/tests/fortran/test_generic_resolution.rs
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
