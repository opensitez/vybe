! vybe-test: fortran/subroutines/derived_type_fixed_array_field_supports_index_read_and_write
! origin: languages/fortran/tests/fortran/test_subroutines.rs
program test
  type :: ode_state
    real :: y(3)
  end type ode_state
  type(ode_state) :: state
  state%y = [1.0, 0.0, 0.0]
  if ((state%y(1)) /= 1) then
    print *, "FAIL: want [1] got [", state%y(1), "]"
    stop 1
end if
  state%y(2) = 4.0
  if ((state%y(2)) /= 4) then
    print *, "FAIL: want [4] got [", state%y(2), "]"
    stop 1
end if
end program test
