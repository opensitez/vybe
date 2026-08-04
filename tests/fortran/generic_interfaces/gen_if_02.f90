! vybe-test: fortran/generic_interfaces/gen_if_02
! origin: languages/fortran/tests/fortran/test_generic_interfaces.rs
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
