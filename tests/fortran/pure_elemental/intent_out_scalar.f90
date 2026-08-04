! vybe-test: fortran/pure_elemental/intent_out_scalar
! origin: languages/fortran/tests/fortran/test_pure_elemental.rs

program test
    integer :: x
    call set_value(x)
    print *, x
contains
    subroutine set_value(n)
        integer, intent(out) :: n
        n = 42
    end subroutine set_value
end program test
