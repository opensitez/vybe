! vybe-test: fortran/array_transfer_between_kinds/array_transfer_between_kinds_character_length_expression
! origin: languages/fortran/tests/fortran/test_array_transfer_between_kinds.rs

program array_transfer_between_kinds_character_length_expression
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 5 ]
    character(len=5) :: token
    integer :: value
    token = transfer(12345, token)
    value = transfer(token, value)
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((len(token)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", len(token), "]"
        stop 1
    end if
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((value) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", value, "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program array_transfer_between_kinds_character_length_expression
