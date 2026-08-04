! vybe-test: fortran/select_case_advanced/case_multiple_values
! origin: languages/fortran/tests/fortran/test_select_case_advanced.rs

program test
    integer :: n = 3
    select case (n)
    case (1, 3, 5, 7, 9)
        print *, 'odd'
    case (2, 4, 6, 8, 10)
        print *, 'even'
    end select
end program test
