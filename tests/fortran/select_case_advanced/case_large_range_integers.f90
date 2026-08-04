! vybe-test: fortran/select_case_advanced/case_large_range_integers
! origin: languages/fortran/tests/fortran/test_select_case_advanced.rs

program test
    integer :: n = 5000
    select case (n)
    case (1:999)
        print *, 'hundreds'
    case (1000:9999)
        print *, 'thousands'
    case (10000:)
        print *, 'ten-thousands+'
    end select
end program test
