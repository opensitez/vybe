! vybe-test: fortran/select_type_rank_extended/fortran_2018_select_rank_runtime
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(2) = [ 1, 0 ]
    call handle([1, 2, 3])
    call inspect(42)
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
contains
    subroutine handle(x)
        integer, intent(in) :: x(..)
        select rank(x)
        rank(1)
                        vybe_check_i = vybe_check_i + 1
            if (vybe_check_i > 2) then
                print *, "FAIL: more than 2 line(s)"
                stop 1
            end if
            if ((1) /= vybe_check_w(vybe_check_i)) then
                print *, "FAIL at ", vybe_check_i, " got [", 1, "]"
                stop 1
            end if
        rank default
                        vybe_check_i = vybe_check_i + 1
            if (vybe_check_i > 2) then
                print *, "FAIL: more than 2 line(s)"
                stop 1
            end if
            if ((0) /= vybe_check_w(vybe_check_i)) then
                print *, "FAIL at ", vybe_check_i, " got [", 0, "]"
                stop 1
            end if
        end select
    end subroutine handle

    subroutine inspect(x)
        integer, intent(in) :: x(..)
        select rank(x)
        rank(0)
                        vybe_check_i = vybe_check_i + 1
            if (vybe_check_i > 2) then
                print *, "FAIL: more than 2 line(s)"
                stop 1
            end if
            if ((0) /= vybe_check_w(vybe_check_i)) then
                print *, "FAIL at ", vybe_check_i, " got [", 0, "]"
                stop 1
            end if
        rank default
                        vybe_check_i = vybe_check_i + 1
            if (vybe_check_i > 2) then
                print *, "FAIL: more than 2 line(s)"
                stop 1
            end if
            if ((-1) /= vybe_check_w(vybe_check_i)) then
                print *, "FAIL at ", vybe_check_i, " got [", -1, "]"
                stop 1
            end if
        end select
    end subroutine inspect
end program t
