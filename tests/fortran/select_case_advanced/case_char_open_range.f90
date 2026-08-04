! vybe-test: fortran/select_case_advanced/case_char_open_range
! origin: languages/fortran/tests/fortran/test_select_case_advanced.rs

program test
    character :: c = 'Z'
    select case (c)
    case ('A':'Z')
        print *, 'uppercase'
    case ('a':'z')
        print *, 'lowercase'
    case default
        print *, 'other'
    end select
end program test
