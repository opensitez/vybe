! vybe-test: fortran/generic_ambiguity/generic_ambiguity_06
! origin: languages/fortran/tests/fortran/test_generic_ambiguity.rs
module m
interface g
module procedure s1
end interface
contains
subroutine s1(i)
integer::i
end
end module m
program p
use m
call g(1)
end program p
