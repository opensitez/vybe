! vybe-test: fortran/select_case_advanced/case_from_logical_merge
! origin: languages/fortran/tests/fortran/test_select_case_advanced.rs

program test
    logical :: flag = .true.
    select case (merge(1, 0, flag))
    case (0)
        print *, 'false'
    case (1)
        print *, 'true'
    end select
end program test
