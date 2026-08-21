! vybe-test: fortran/select_type_rank_extended/select_rank_explicit_rank3_branch_runtime
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs

program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(3) = [ 2, 3, 4 ]
    call inspect(reshape([(i, i = 1, 24)], [2, 3, 4]))
if (vybe_check_i /= 3) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 3"
    stop 1
end if
contains
    subroutine inspect(x)
        integer, intent(in) :: x(..)
        select rank(x)
        rank(3)
                        vybe_check_i = vybe_check_i + 1
            if (vybe_check_i > 3) then
                print *, "FAIL: more than 3 line(s)"
                stop 1
            end if
            if ((size(x, 1)) /= vybe_check_w(vybe_check_i)) then
                print *, "FAIL at ", vybe_check_i, " got [", size(x, 1), "]"
                stop 1
            end if
                        vybe_check_i = vybe_check_i + 1
            if (vybe_check_i > 3) then
                print *, "FAIL: more than 3 line(s)"
                stop 1
            end if
            if ((size(x, 2)) /= vybe_check_w(vybe_check_i)) then
                print *, "FAIL at ", vybe_check_i, " got [", size(x, 2), "]"
                stop 1
            end if
                        vybe_check_i = vybe_check_i + 1
            if (vybe_check_i > 3) then
                print *, "FAIL: more than 3 line(s)"
                stop 1
            end if
            if ((size(x, 3)) /= vybe_check_w(vybe_check_i)) then
                print *, "FAIL at ", vybe_check_i, " got [", size(x, 3), "]"
                stop 1
            end if
        rank default
                        vybe_check_i = vybe_check_i + 1
            if (vybe_check_i > 3) then
                print *, "FAIL: more than 3 line(s)"
                stop 1
            end if
            if ((rank(x)) /= vybe_check_w(vybe_check_i)) then
                print *, "FAIL at ", vybe_check_i, " got [", rank(x), "]"
                stop 1
            end if
        end select
    end subroutine inspect
end program t
