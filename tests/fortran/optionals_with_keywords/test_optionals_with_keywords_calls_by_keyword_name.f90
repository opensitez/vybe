! vybe-test: fortran/optionals_with_keywords/test_optionals_with_keywords_calls_by_keyword_name
! origin: languages/fortran/tests/fortran/test_optionals_with_keywords.rs

program test_optionals_with_keywords
    call configure(scale=3, mode=2)
    call configure(scale=3)

contains
    subroutine configure(scale, mode)
        integer, intent(in) :: scale
        integer, optional, intent(in) :: mode
        if (present(mode)) print *, scale + mode
        if (.not. present(mode)) print *, scale
    end subroutine
end program test_optionals_with_keywords
