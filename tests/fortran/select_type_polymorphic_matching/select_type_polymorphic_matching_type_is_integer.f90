! vybe-test: fortran/select_type_polymorphic_matching/select_type_polymorphic_matching_type_is_integer
! origin: languages/fortran/tests/fortran/test_select_type_polymorphic_matching.rs

program select_type_polymorphic_matching_type_is_integer
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 4 ]
    class(*), allocatable :: value
    allocate(integer :: value)
    value = 4
    select type (value)
    type is (integer)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((value) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", value, "]"
            stop 1
        end if
    class default
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((-1) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", -1, "]"
            stop 1
        end if
    end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program select_type_polymorphic_matching_type_is_integer
