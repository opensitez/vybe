! vybe-test: fortran/specification_part/spec_import_22
! origin: languages/fortran/tests/fortran/test_specification_part.rs
module m
implicit none
contains
 subroutine s()
  import
 end subroutine s
end module m
