! vybe-test: fortran/subroutines/interface_procedure_member_aliases_signature
! origin: languages/fortran/tests/fortran/test_subroutines.rs
module m
  implicit none
  abstract interface
    integer function unary(x) result(v)
      integer, intent(in) :: x
      integer :: v
    end function unary
    procedure(unary) :: op
  end interface
contains
  subroutine step(rhs)
    procedure(op) :: rhs
    if ((rhs(3)) /= 6) then
    print *, "FAIL: want [6] got [", rhs(3), "]"
    stop 1
end if
  end subroutine step
end module m

program test
  use m
  call step(double_it)
contains
  integer function double_it(x) result(v)
    integer, intent(in) :: x
    integer :: v
    v = x * 2
  end function double_it
end program test
