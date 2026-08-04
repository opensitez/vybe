! vybe-test: fortran/interfaces/if_generic_two_05
! origin: languages/fortran/tests/fortran/test_interfaces.rs
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
