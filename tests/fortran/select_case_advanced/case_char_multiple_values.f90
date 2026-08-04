! vybe-test: fortran/select_case_advanced/case_char_multiple_values
! origin: languages/fortran/tests/fortran/test_select_case_advanced.rs

program test
    character :: c = 'e'
    select case (c)
    case ('a', 'e', 'i', 'o', 'u')
        print *, 'vowel'
    case default
        print *, 'consonant'
    end select
end program test
