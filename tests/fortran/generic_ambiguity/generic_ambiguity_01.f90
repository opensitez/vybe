! vybe-test: fortran/generic_ambiguity/generic_ambiguity_01
! origin: languages/fortran/tests/fortran/test_generic_ambiguity.rs
module m
interface g
module procedure si,sr
end interface
contains
subroutine si(i)
integer::i
end
subroutine sr(r)
real::r
end
end module m
