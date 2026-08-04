! vybe-test: fortran/array_section_strictness_and_errors/array_section_strictness_and_errors_zero_stride_is_rejected_by_parser
! origin: languages/fortran/tests/fortran/test_array_section_strictness_and_errors.rs
program array_section_strictness_and_errors_zero_stride_is_rejected_by_parser
integer :: values(1:5)
print *, values(1:5:0)
end program
