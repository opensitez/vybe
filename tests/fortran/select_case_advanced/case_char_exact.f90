! vybe-test: fortran/select_case_advanced/case_char_exact
! origin: languages/fortran/tests/fortran/test_select_case_advanced.rs

program test
    character :: c = 'b'
    select case (c)
    case ('a')
        print *, 'a'
    case ('b')
        print *, 'b'
    case ('c')
        print *, 'c'
    end select
end program test
