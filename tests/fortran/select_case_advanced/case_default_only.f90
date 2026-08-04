! vybe-test: fortran/select_case_advanced/case_default_only
! origin: languages/fortran/tests/fortran/test_select_case_advanced.rs

program test
    integer :: n = 42
    select case (n)
    case default
        print *, 'default'
    end select
end program test
