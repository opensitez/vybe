! vybe-test: fortran/select_type_polymorphic_matching/select_type_polymorphic_matching_character_dispatch
! origin: languages/fortran/tests/fortran/test_select_type_polymorphic_matching.rs

program select_type_polymorphic_matching_character_dispatch
integer :: vybe_check_i = 0
character(len=4) :: vybe_check_w(1) = [ "data" ]
    class(*), allocatable :: value
    allocate(character(len=4) :: value)
    value = 'data'
    select type (value)
    type is (character(len=*))
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim(trim(value)) /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", trim(value), "]"
            stop 1
        end if
    class default
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('nope') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'nope', "]"
            stop 1
        end if
    end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program select_type_polymorphic_matching_character_dispatch
