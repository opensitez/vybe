! vybe-test: fortran/fortran2018_extended/select_rank_catches_rank1_vector
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs

program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 3 ]
    call tag([10, 20, 30])
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
contains
    subroutine tag(x)
        integer, intent(in) :: x(..)
        select rank(x)
        rank(1)
                        vybe_check_i = vybe_check_i + 1
            if (vybe_check_i > 1) then
                print *, "FAIL: more than 1 line(s)"
                stop 1
            end if
            if ((size(x)) /= vybe_check_w(vybe_check_i)) then
                print *, "FAIL at ", vybe_check_i, " got [", size(x), "]"
                stop 1
            end if
        rank default
                        vybe_check_i = vybe_check_i + 1
            if (vybe_check_i > 1) then
                print *, "FAIL: more than 1 line(s)"
                stop 1
            end if
            if ((0) /= vybe_check_w(vybe_check_i)) then
                print *, "FAIL at ", vybe_check_i, " got [", 0, "]"
                stop 1
            end if
        end select
    end subroutine tag
end program t
