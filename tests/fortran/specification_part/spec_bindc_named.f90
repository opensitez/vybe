! vybe-test: fortran/specification_part/spec_bindc_named
! origin: languages/fortran/tests/fortran/test_specification_part.rs
subroutine s() bind(c, name='c_entry')
implicit none
end subroutine s
