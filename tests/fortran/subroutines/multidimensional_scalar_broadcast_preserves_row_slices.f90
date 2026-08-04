! vybe-test: fortran/subroutines/multidimensional_scalar_broadcast_preserves_row_slices
! origin: languages/fortran/tests/fortran/test_subroutines.rs
program test
  integer :: i
  real :: trajectory(4, 5)
  trajectory = 0.0
  trajectory(2, :) = [3.0, -2.0, 7.0, 1.5, 4.0]
  i = 1
  if (trim(trajectory(2)) /= "3,-2,7,1.5,4") then
    print *, "FAIL: want [3,-2,7,1.5,4] got [", trajectory(2), "]"
    stop 1
end if
  if ((trajectory(2, 1)) /= 3) then
    print *, "FAIL: want [3] got [", trajectory(2, 1), "]"
    stop 1
end if
  if ((minval(trajectory(i + 1, 1:5))) /= -2) then
    print *, "FAIL: want [-2] got [", minval(trajectory(i + 1, 1:5)), "]"
    stop 1
end if
  if ((maxval(trajectory(i + 1, 1:5))) /= 7) then
    print *, "FAIL: want [7] got [", maxval(trajectory(i + 1, 1:5)), "]"
    stop 1
end if
end program test
