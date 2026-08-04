! vybe-test: fortran/select_case_advanced/case_range_boundary_exact
! origin: languages/fortran/tests/fortran/test_select_case_advanced.rs

program test
    integer :: n
    do n = 4, 6
        select case (n)
        case (:4)
            print *, 'le 4'
        case (5)
            print *, 'eq 5'
        case (6:)
            print *, 'ge 6'
        end select
    end do
end program test
