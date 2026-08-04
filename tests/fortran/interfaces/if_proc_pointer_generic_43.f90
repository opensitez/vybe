! vybe-test: fortran/interfaces/if_proc_pointer_generic_43
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
abstract interface
subroutine fn(x, y)
integer, intent(in) :: x
real, intent(in) :: y
end subroutine fn
end interface
end module m
