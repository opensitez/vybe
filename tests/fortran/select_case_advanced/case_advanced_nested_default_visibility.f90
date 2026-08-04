! vybe-test: fortran/select_case_advanced/case_advanced_nested_default_visibility
! origin: languages/fortran/tests/fortran/test_select_case_advanced.rs

program test
integer :: vybe_check_i = 0
character(len=13) :: vybe_check_w(2) = [ "inner-default", "outer-default" ]
    integer :: x, y
    x = 3
    y = 2
    select case (x)
    case (1, 2)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 2) then
            print *, "FAIL: more than 2 line(s)"
            stop 1
        end if
        if (trim('outer-low') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'outer-low', "]"
            stop 1
        end if
    case default
        select case (y)
        case (1)
                        vybe_check_i = vybe_check_i + 1
            if (vybe_check_i > 2) then
                print *, "FAIL: more than 2 line(s)"
                stop 1
            end if
            if (trim('inner-one') /= trim(vybe_check_w(vybe_check_i))) then
                print *, "FAIL at ", vybe_check_i, " got [", 'inner-one', "]"
                stop 1
            end if
        case default
                        vybe_check_i = vybe_check_i + 1
            if (vybe_check_i > 2) then
                print *, "FAIL: more than 2 line(s)"
                stop 1
            end if
            if (trim('inner-default') /= trim(vybe_check_w(vybe_check_i))) then
                print *, "FAIL at ", vybe_check_i, " got [", 'inner-default', "]"
                stop 1
            end if
        end select
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 2) then
            print *, "FAIL: more than 2 line(s)"
            stop 1
        end if
        if (trim('outer-default') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'outer-default', "]"
            stop 1
        end if
    end select
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
end program test
