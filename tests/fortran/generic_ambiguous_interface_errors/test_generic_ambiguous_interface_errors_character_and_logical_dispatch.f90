! vybe-test: fortran/generic_ambiguous_interface_errors/test_generic_ambiguous_interface_errors_character_and_logical_dispatch
! origin: languages/fortran/tests/fortran/test_generic_ambiguous_interface_errors.rs

program test_generic_ambiguous_interface_errors
    if ((magnitude('abc')) /= 3) then
    print *, "FAIL: want [3] got [", magnitude('abc'), "]"
    stop 1
end if
    if ((magnitude(.true.)) /= 1) then
    print *, "FAIL: want [1] got [", magnitude(.true.), "]"
    stop 1
end if

contains
    interface magnitude
        module procedure abs_char
        module procedure abs_log
    end interface

    integer function abs_char(v)
        character(len=*), intent(in) :: v
        abs_char = len_trim(v)
    end function

    integer function abs_log(v)
        logical, intent(in) :: v
        if (v) then
            abs_log = 1
        else
            abs_log = 0
        end if
    end function
end program test_generic_ambiguous_interface_errors
