! vybe-test: fortran/select_case_advanced/case_no_match_no_default
! origin: languages/fortran/tests/fortran/test_select_case_advanced.rs

program test
    integer :: n = 99
    select case (n)
    case (1)
        print *, 'one'
    case (2)
        print *, 'two'
    end select
    print *, 'after select'
end program test
