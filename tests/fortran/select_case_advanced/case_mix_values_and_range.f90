! vybe-test: fortran/select_case_advanced/case_mix_values_and_range
! origin: languages/fortran/tests/fortran/test_select_case_advanced.rs

program test
    integer :: n = 0
    select case (n)
    case (0, 1, 2)
        print *, 'small'
    case (3:10)
        print *, 'medium'
    case (11:)
        print *, 'large'
    end select
end program test
