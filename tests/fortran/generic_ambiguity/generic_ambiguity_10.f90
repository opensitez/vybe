! vybe-test: fortran/generic_ambiguity/generic_ambiguity_10
! origin: languages/fortran/tests/fortran/test_generic_ambiguity.rs
module m
interface g
module procedure s1,s2
end interface
contains
subroutine s1(c)
character(len=*)::c
end
subroutine s2(l)
logical::l
end
end module m
