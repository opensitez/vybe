! vybe-test: fortran/select_case_advanced/case_on_function_result
! origin: languages/fortran/tests/fortran/test_select_case_advanced.rs

program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    select case (sum(a))
    case (:10)
        print *, 'small'
    case (11:20)
        print *, 'medium'
    case (21:)
        print *, 'large'
    end select
end program test
