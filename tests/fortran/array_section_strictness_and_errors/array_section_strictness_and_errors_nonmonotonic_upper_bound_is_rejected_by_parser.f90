! vybe-test: fortran/array_section_strictness_and_errors/array_section_strictness_and_errors_nonmonotonic_upper_bound_is_rejected_by_parser
! origin: languages/fortran/tests/fortran/test_array_section_strictness_and_errors.rs
program array_section_strictness_and_errors_nonmonotonic_upper_bound_is_rejected_by_parser
integer :: values(1:5)
values = (/1, 2, 3, 4, 5/)
if ((size(values(5:1:-2))) /= 3) then
    print *, "FAIL: want [3] got [", size(values(5:1:-2)), "]"
    stop 1
end if
if ((values(5:1:-2)(2)) /= 3) then
    print *, "FAIL: want [3] got [", values(5:1:-2)(2), "]"
    stop 1
end if
end program
