! vybe-test: fortran/select_case_advanced/case_on_expression
! origin: languages/fortran/tests/fortran/test_select_case_advanced.rs

program test
    integer :: x = 5, y = 3
    select case (x + y)
    case (:7)
        print *, 'small sum'
    case (8:10)
        print *, 'medium sum'
    case (11:)
        print *, 'large sum'
    end select
end program test
