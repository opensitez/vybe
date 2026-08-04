! vybe-test: fortran/generic_interfaces/gen_if_18
! origin: languages/fortran/tests/fortran/test_generic_interfaces.rs
module m
interface g
module procedure ss
end interface
contains
subroutine ss()
print *,1
end
end module m
