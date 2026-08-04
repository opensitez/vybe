! vybe-test: fortran/select_case_advanced/case_multiple_values_some_match
! origin: languages/fortran/tests/fortran/test_select_case_advanced.rs

program test
    integer :: i
    do i = 1, 6
        select case (i)
        case (1, 2, 6)
            print *, 'match'
        case default
            print *, 'no'
        end select
    end do
end program test
