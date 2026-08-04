! vybe-test: fortran/legacy/error_stop
! origin: languages/fortran/tests/fortran/test_legacy.rs
program t
  logical :: ok = .true.
  if (.not. ok) error stop 'fatal'
  print *, 'fine'
end program t
