! vybe-test: fortran/select_case_advanced/nested_select_in_loop
! origin: languages/fortran/tests/fortran/test_select_case_advanced.rs

program test
    integer :: i, j
    do i = 1, 3
        select case (i)
        case (1)
            do j = 1, 2
                select case (j)
                case (1)
                    print *, '1,1'
                case (2)
                    print *, '1,2'
                end select
            end do
        case (2:3)
            print *, 'i=', i
        end select
    end do
end program test
