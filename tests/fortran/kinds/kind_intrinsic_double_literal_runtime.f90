! vybe-test: fortran/kinds/kind_intrinsic_double_literal_runtime
! origin: languages/fortran/tests/fortran/test_kinds.rs
program t
  integer, parameter :: dp = kind(1.0d0)
  if ((dp) /= 8) then
    print *, "FAIL: want [8] got [", dp, "]"
    stop 1
end if
end program t
