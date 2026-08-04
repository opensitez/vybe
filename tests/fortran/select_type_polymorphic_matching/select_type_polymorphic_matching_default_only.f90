! vybe-test: fortran/select_type_polymorphic_matching/select_type_polymorphic_matching_default_only
! origin: languages/fortran/tests/fortran/test_select_type_polymorphic_matching.rs

program select_type_polymorphic_matching_default_only
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 2 ]
    class(*), allocatable :: marker
    allocate(integer(2) :: marker)
    select type (marker)
    class is (logical)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((0) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", 0, "]"
            stop 1
        end if
    class default
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((size(marker)) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", size(marker), "]"
            stop 1
        end if
    end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program select_type_polymorphic_matching_default_only
