! vybe-test: fortran/array_section_strictness_and_errors/array_section_strictness_and_errors_empty_reverse_triplet_has_zero_size
! origin: languages/fortran/tests/fortran/test_array_section_strictness_and_errors.rs

program array_section_strictness_and_errors_empty_reverse_triplet_has_zero_size
    integer :: values(1:5)
    values = (/1, 2, 3, 4, 5/)
    if ((size(values(1:5:-1))) /= 0) then
    print *, "FAIL: want [0] got [", size(values(1:5:-1)), "]"
    stop 1
end if
end program array_section_strictness_and_errors_empty_reverse_triplet_has_zero_size
