! vybe-test: fortran/program_units/program_generic_resolution_21
! origin: languages/fortran/tests/fortran/test_program_units.rs
module m
interface g
module procedure si, sr
end interface
contains
subroutine si(i)
integer :: i
end subroutine si
subroutine sr(r)
real :: r
end subroutine sr
end module m
