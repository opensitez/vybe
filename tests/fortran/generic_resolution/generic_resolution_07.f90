! vybe-test: fortran/generic_resolution/generic_resolution_07
! origin: languages/fortran/tests/fortran/test_generic_resolution.rs
module m
interface g
module procedure ss
end interface
contains
subroutine ss(s)
character(len=*)::s
end
end module m
