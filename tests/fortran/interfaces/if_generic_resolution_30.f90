! vybe-test: fortran/interfaces/if_generic_resolution_30
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
interface g
module procedure s1,s2
end interface
contains
subroutine s1(i)
integer::i
end
subroutine s2(r)
real::r
end
end module m
