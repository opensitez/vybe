! vybe-test: fortran/select_case_advanced/nested_select_case
! origin: languages/fortran/tests/fortran/test_select_case_advanced.rs

program test
    integer :: i = 2, j = 3
    select case (i)
    case (1)
        print *, 'i=1'
    case (2)
        select case (j)
        case (1:2)
            print *, 'i=2, j small'
        case (3:)
            print *, 'i=2, j large'
        end select
    case default
        print *, 'other'
    end select
end program test
