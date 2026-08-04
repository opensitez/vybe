! vybe-test: fortran/io_advanced/read_member_targets
! origin: languages/fortran/tests/fortran/test_io_advanced.rs
program t
  type :: field_t
    real :: data(4)
  end type field_t
  type :: state_t
    real :: time
    type(field_t) :: h
  end type state_t
  type(state_t) :: state
  integer :: unit
  read(unit) state%time
  read(unit) state%h%data
end program t
