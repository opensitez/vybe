! vybe-test: fortran/specification_part/spec_contiguous_28
! origin: languages/fortran/tests/fortran/test_specification_part.rs
subroutine s(a)
implicit none
integer, contiguous :: a(:)
end subroutine s
