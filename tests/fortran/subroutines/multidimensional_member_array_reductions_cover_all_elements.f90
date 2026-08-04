! vybe-test: fortran/subroutines/multidimensional_member_array_reductions_cover_all_elements
! origin: languages/fortran/tests/fortran/test_subroutines.rs
program test
  type :: field2d
    real, allocatable :: data(:,:)
  contains
    procedure :: init => field_init
  end type
  type(field2d) :: h, u
  call h%init(2,2)
  call u%init(2,2)
  h%data = 3.0
  u%data = 4.0
  if ((size(h%data)) /= 4) then
    print *, "FAIL: want [4] got [", size(h%data), "]"
    stop 1
end if
  if ((maxval(h%data * u%data)) /= 12) then
    print *, "FAIL: want [12] got [", maxval(h%data * u%data), "]"
    stop 1
end if
  if ((sum(h%data * u%data)) /= 48) then
    print *, "FAIL: want [48] got [", sum(h%data * u%data), "]"
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
