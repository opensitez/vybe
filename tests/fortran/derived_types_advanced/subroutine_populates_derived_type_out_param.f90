! vybe-test: fortran/derived_types_advanced/subroutine_populates_derived_type_out_param
! origin: languages/fortran/tests/fortran/test_derived_types_advanced.rs

program test
    type :: Counter
        integer :: n = 0
    end type Counter
    type(Counter) :: c
    call fill(c)
    if ((c%n) /= 7) then
    print *, "FAIL: want [7] got [", c%n, "]"
    stop 1
end if
contains
    subroutine fill(counter)
        type(Counter), intent(out) :: counter
        counter%n = 7
    end subroutine fill
end program test
