! vybe-test: fortran/generic_ambiguity/generic_ambiguity_08
! origin: languages/fortran/tests/fortran/test_generic_ambiguity.rs
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
program p
use m
call g(1)
call g(1.0)
end program p
