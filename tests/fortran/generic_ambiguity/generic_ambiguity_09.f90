! vybe-test: fortran/generic_ambiguity/generic_ambiguity_09
! origin: languages/fortran/tests/fortran/test_generic_ambiguity.rs
module m
interface operator(-)
module procedure subi,subr
end interface
contains
integer function subi(a,b)
integer::a,b
subi=a-b
end
real function subr(a,b)
real::a,b
subr=a-b
end
end module m
