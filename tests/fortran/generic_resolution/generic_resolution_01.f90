! vybe-test: fortran/generic_resolution/generic_resolution_01
! origin: languages/fortran/tests/fortran/test_generic_resolution.rs
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
