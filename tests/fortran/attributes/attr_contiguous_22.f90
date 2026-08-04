! vybe-test: fortran/attributes/attr_contiguous_22
! origin: languages/fortran/tests/fortran/test_attributes.rs
subroutine s(a)
real,contiguous::a(:)
end subroutine s
