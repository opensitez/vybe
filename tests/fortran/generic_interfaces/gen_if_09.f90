! vybe-test: fortran/generic_interfaces/gen_if_09
! origin: languages/fortran/tests/fortran/test_generic_interfaces.rs
module m
interface write(formatted)
module procedure wf
end interface
contains
subroutine wf()
end subroutine wf
end module m
