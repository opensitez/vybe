! vybe-test: fortran/subroutines/nested_allocatable_member_index_assignment_writes_back_receiver
! origin: languages/fortran/tests/fortran/test_subroutines.rs
program test
  type :: field2d
    real, allocatable :: data(:,:)
  contains
    procedure :: init => field_init
  end type field2d
  type :: state_t
    type(field2d) :: h
  end type state_t
  type(state_t) :: state
  call state%h%init(3, 4)
  state%h%data(2, 3) = 42.0
  state%h%data(1, 1) = -7.5
  if ((state%h%data(2, 3)) /= 42) then
    print *, "FAIL: want [42] got [", state%h%data(2, 3), "]"
    stop 1
end if
  if (abs((state%h%data(1, 1)) - -7.5) > 1.0e-6) then
    print *, "FAIL: want [-7.5] got [", state%h%data(1, 1), "]"
    stop 1
end if
contains
  subroutine field_init(self, nx, ny)
    class(field2d), intent(inout) :: self
    integer, intent(in) :: nx, ny
    allocate(self%data(nx, ny))
    self%data = 0.0
  end subroutine field_init
end program test
