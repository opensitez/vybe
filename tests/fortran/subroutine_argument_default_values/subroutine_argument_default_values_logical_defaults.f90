! vybe-test: fortran/subroutine_argument_default_values/subroutine_argument_default_values_logical_defaults
! origin: languages/fortran/tests/fortran/test_subroutine_argument_default_values.rs

program subroutine_argument_default_values_logical_defaults
    logical :: a
    logical :: b
    call toggle(a)
    call toggle(b, .false.)
    if ((a) .neqv. .true.) then
    print *, "FAIL: want [True] got [", a, "]"
    stop 1
end if
    if ((b) .neqv. .false.) then
    print *, "FAIL: want [False] got [", b, "]"
    stop 1
end if
contains
    subroutine toggle(value, active)
        logical, intent(out) :: value
        logical, intent(in), optional :: active
        if (present(active)) then
            value = active
        else
            value = .true.
        end if
    end subroutine toggle
end program subroutine_argument_default_values_logical_defaults
