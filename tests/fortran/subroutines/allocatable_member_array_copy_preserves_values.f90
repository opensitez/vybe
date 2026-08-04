! vybe-test: fortran/subroutines/allocatable_member_array_copy_preserves_values
! origin: languages/fortran/tests/fortran/test_subroutines.rs
program test
  type :: box
    integer :: n
    real, allocatable :: y(:)
  contains
    procedure :: init => box_init
    procedure :: copy => box_copy
  end type
  type(box) :: a, b
  call a%init(3)
  a%y = [1.0, 2.0, 3.0]
  call b%copy(a)
  if ((sum(b%y)) /= 6) then
    print *, "FAIL: want [6] got [", sum(b%y), "]"
    stop 1
end if
contains
  subroutine box_init(self, n)
    class(box), intent(inout) :: self
    integer, intent(in) :: n
    self%n = n
    allocate(self%y(n))
    self%y = 0.0
  end subroutine box_init
  subroutine box_copy(self, other)
    class(box), intent(inout) :: self
    type(box), intent(in) :: other
    self%n = other%n
    if (allocated(self%y)) deallocate(self%y)
    allocate(self%y(other%n))
    self%y = other%y
  end subroutine box_copy
end program test
