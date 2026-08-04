! vybe-test: fortran/select_case_advanced/case_advanced_nested_no_match_preserves_outer_default
! origin: languages/fortran/tests/fortran/test_select_case_advanced.rs

program test
integer :: vybe_check_i = 0
character(len=17) :: vybe_check_w(2) = [ "inner-default", "outer-after-inner" ]
    integer :: i, j
    i = 10
    j = 0
    select case (i)
    case (10)
        select case (j)
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
        if (trim('outer-after-inner') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'outer-after-inner', "]"
            stop 1
        end if
    case default
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
