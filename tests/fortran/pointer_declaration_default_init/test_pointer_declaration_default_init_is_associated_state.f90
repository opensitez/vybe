! vybe-test: fortran/pointer_declaration_default_init/test_pointer_declaration_default_init_is_associated_state
! origin: languages/fortran/tests/fortran/test_pointer_declaration_default_init.rs

program test_pointer_declaration_default_init
integer :: vybe_check_i = 0
integer :: vybe_check_w(2) = [ 1, 9 ]
    integer, target :: storage
    integer, pointer :: p => null()
    if (associated(p)) then
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 2) then
            print *, "FAIL: more than 2 line(s)"
            stop 1
        end if
        if ((0) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", 0, "]"
            stop 1
        end if
    else
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 2) then
            print *, "FAIL: more than 2 line(s)"
            stop 1
        end if
        if ((1) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", 1, "]"
            stop 1
        end if
    end if
    p => storage
    storage = 9
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 2) then
        print *, "FAIL: more than 2 line(s)"
        stop 1
    end if
    if ((p) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", p, "]"
        stop 1
    end if
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
end program test_pointer_declaration_default_init
