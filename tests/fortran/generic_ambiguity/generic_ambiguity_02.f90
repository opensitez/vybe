! vybe-test: fortran/generic_ambiguity/generic_ambiguity_02
! origin: languages/fortran/tests/fortran/test_generic_ambiguity.rs
module m
interface g
module procedure s1,s2
end interface
contains
subroutine s1(i)
integer::i
end
subroutine s2(j)
integer::j
end
end module m
