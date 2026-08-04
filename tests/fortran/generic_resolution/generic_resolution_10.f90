! vybe-test: fortran/generic_resolution/generic_resolution_10
! origin: languages/fortran/tests/fortran/test_generic_resolution.rs
module m
interface g
module procedure li
end interface
contains
subroutine li(l)
logical::l
end
end module m
