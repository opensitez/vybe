! vybe-test: fortran/interfaces/if_generic_result_type_46
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
interface g
module procedure gi, gr
end interface
contains
integer function gi(x)
integer, intent(in) :: x
gi = x
end function gi
real function gr(x)
real, intent(in) :: x
gr = x
end function gr
end module m
