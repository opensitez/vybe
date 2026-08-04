! vybe-test: fortran/select_case_advanced/case_on_mod_result
! origin: languages/fortran/tests/fortran/test_select_case_advanced.rs

program test
    integer :: i
    do i = 1, 6
        select case (mod(i, 3))
        case (0)
            print *, i, 'div by 3'
        case (1)
            print *, i, 'rem 1'
        case (2)
            print *, i, 'rem 2'
        end select
    end do
end program test
