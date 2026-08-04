! vybe-test: fortran/select_case_advanced/case_open_both_ends
! origin: languages/fortran/tests/fortran/test_select_case_advanced.rs

program test
    integer :: n = 50
    select case (n)
    case (:9)
        print *, 'single digit'
    case (10:99)
        print *, 'double digit'
    case (100:)
        print *, 'triple digit or more'
    end select
end program test
