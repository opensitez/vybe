! vybe-test: fortran/generic_interfaces/gen_if_08
! origin: languages/fortran/tests/fortran/test_generic_interfaces.rs
module m
interface g
procedure f
end interface
contains
subroutine f()
end subroutine f
end module m
