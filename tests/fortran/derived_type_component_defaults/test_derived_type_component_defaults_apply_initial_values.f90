! vybe-test: fortran/derived_type_component_defaults/test_derived_type_component_defaults_apply_initial_values
! origin: languages/fortran/tests/fortran/test_derived_type_component_defaults.rs

program test_derived_type_component_defaults
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 3 ]
    type :: item
        integer :: x = 3
        logical :: enabled = .true.
    end type

    type(item) :: a
    if (a%enabled) then
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((a%x) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", a%x, "]"
            stop 1
        end if
    else
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((-1) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", -1, "]"
            stop 1
        end if
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test_derived_type_component_defaults
