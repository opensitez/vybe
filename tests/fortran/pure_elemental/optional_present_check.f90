! vybe-test: fortran/pure_elemental/optional_present_check
! origin: languages/fortran/tests/fortran/test_pure_elemental.rs

program test
    call maybe(3)
contains
    subroutine maybe(n, extra)
        integer, intent(in) :: n
        integer, intent(in), optional :: extra
        integer :: total
        total = n
        if (present(extra)) total = total + extra
        print *, total
    end subroutine maybe
end program test
