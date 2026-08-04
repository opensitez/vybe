! vybe-test: fortran/array_section_strictness_and_errors/array_section_strictness_and_errors_nonmonotonic_upper_bound_is_rejected_by_parser
! origin: languages/fortran/tests/fortran/test_array_section_strictness_and_errors.rs
program array_section_strictness_and_errors_nonmonotonic_upper_bound_is_rejected_by_parser
integer :: values(1:5)
print *, values(1:3:2:1)
end program
