! vybe-test: fortran/select_type_polymorphic_matching/select_type_polymorphic_matching_logical_dispatch
! origin: languages/fortran/tests/fortran/test_select_type_polymorphic_matching.rs

program select_type_polymorphic_matching_logical_dispatch
integer :: vybe_check_i = 0
logical :: vybe_check_w(1) = [ .true. ]
    class(*), allocatable :: value
    allocate(logical :: value)
    value = .true.
    select type (value)
    type is (logical)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((value) .neqv. vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", value, "]"
            stop 1
        end if
    class default
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((.false.) .neqv. vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", .false., "]"
            stop 1
        end if
    end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program select_type_polymorphic_matching_logical_dispatch
