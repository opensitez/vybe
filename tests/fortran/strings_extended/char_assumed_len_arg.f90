! vybe-test: fortran/strings_extended/char_assumed_len_arg
! origin: languages/fortran/tests/fortran/test_strings_extended.rs

program test
integer :: vybe_check_i = 0
character(len=5) :: vybe_check_w(2) = [ "alice", "bob" ]
    call print_it('alice')
    call print_it('bob')
contains
    subroutine print_it(s)
        character(len=*), intent(in) :: s
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 2) then
            print *, "FAIL: more than 2 line(s)"
            stop 1
        end if
        if (trim(trim(s)) /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", trim(s), "]"
            stop 1
        end if
    end subroutine
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
end program test
