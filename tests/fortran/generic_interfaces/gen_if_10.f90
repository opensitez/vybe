! vybe-test: fortran/generic_interfaces/gen_if_10
! origin: languages/fortran/tests/fortran/test_generic_interfaces.rs
module m
interface read(formatted)
module procedure rf
end interface
contains
subroutine rf()
end subroutine rf
end module m
