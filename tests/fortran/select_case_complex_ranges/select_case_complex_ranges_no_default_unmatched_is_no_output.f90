! vybe-test: fortran/select_case_complex_ranges/select_case_complex_ranges_no_default_unmatched_is_no_output
! origin: languages/fortran/tests/fortran/test_select_case_complex_ranges.rs

program select_case_complex_ranges_no_default_unmatched_is_no_output
    integer :: n
    n = 0
    select case (n)
    case (1:10)
        print *, 'low'
    case (11:20)
        print *, 'mid'
    end select
end program select_case_complex_ranges_no_default_unmatched_is_no_output
