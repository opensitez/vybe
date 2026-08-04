! vybe-test: fortran/subroutines/procedure_dummy_call_uses_abstract_interface_signature
! origin: languages/fortran/tests/fortran/test_subroutines.rs
module m
  implicit none
  abstract interface
    integer function rhs_func(t) result(v)
      integer, intent(in) :: t
      integer :: v
    end function rhs_func
  end interface
contains
  subroutine step(rhs)
    procedure(rhs_func) :: rhs
    if ((rhs(2)) /= 4) then
    print *, "FAIL: want [4] got [", rhs(2), "]"
    stop 1
end if
  end subroutine step
end module m

program test
  use m
  call step(rhs1)
contains
  integer function rhs1(t) result(v)
    integer, intent(in) :: t
    integer :: v
    v = t * 2
  end function rhs1
end program test
